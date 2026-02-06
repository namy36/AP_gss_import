#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from __future__ import annotations

import ctypes
import datetime as dt
import hashlib
import os
import re
import shutil
import subprocess
import sys
import time
import traceback
from datetime import datetime, timedelta
from pathlib import Path

import cv2
import gspread
import numpy as np
import pandas as pd
import pyautogui as ag
import pygetwindow as gw
import pyperclip
from dateutil.relativedelta import relativedelta
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from gspread.exceptions import APIError, WorksheetNotFound


class MyError(Exception):  # ユーザー定義例外
    pass


# # ウィンドウ名チェック（ウィンドウ名,タイムアウト時間）
# def app_check(app_name, timeout):
#     start = time.time()
#     while time.time() - start <= 60 * timeout:  # タイムアウト
#         app_window = gw.getWindowsWithTitle(app_name)
#         if app_window:  # リストが空でない
#             win = app_window[0]
#             win.activate()  # アクティブ化
#             # ウィンドウタイトルをチェック
#             if '名前を付けて保存' not in win.title:
#                 win.maximize()
#             ime_off()
#             break
#         else:
#             time.sleep(1)
#     if time.time() - start >= 60 * timeout:  # タイムアウトした場合
#         raise MyError('ERROR:' + app_name + 'ウィンドウ取得失敗')


# --- Win32 定義 ---
user32 = ctypes.windll.user32
GWL_STYLE = -16
WS_MAXIMIZEBOX = 0x00010000
WS_THICKFRAME = 0x00040000
SW_MAXIMIZE = 3


def _can_maximize(win):
    hwnd = getattr(win, '_hWnd', None)
    if not hwnd:
        return False
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    return bool(style & WS_MAXIMIZEBOX) and bool(style & WS_THICKFRAME)


def _pick_best(cands, target):
    # 完全一致 > "target (" で始まる（例: 出力 (リモート)）> 先頭
    for w in cands:
        if w.title.strip() == target:
            return w
    for w in cands:
        if w.title.strip().startswith(target + ' ('):
            return w
    return cands[0]


# ウィンドウ名チェック（ウィンドウ名, タイムアウト時間[分]）
def app_check(app_name, timeout):
    start = time.time()
    while time.time() - start <= 60 * timeout:  # タイムアウト
        app_window = gw.getWindowsWithTitle(app_name)
        if app_window:  # リストが空でない
            win = _pick_best(app_window, app_name)

            # 最小化→復帰
            if win.isMinimized:
                try:
                    win.restore()
                    time.sleep(0.05)
                except Exception:
                    pass

            # アクティブ化
            try:
                win.activate()
                time.sleep(0.1)
            except Exception:
                pass

            # 「名前を付けて保存」以外は最大化可能なら最大化
            if '名前を付けて保存' not in win.title and _can_maximize(win):
                try:
                    user32.ShowWindow(getattr(win, '_hWnd', None), SW_MAXIMIZE)  # 安定版
                except Exception:
                    try:
                        win.maximize()  # フォールバック
                    except Exception:
                        pass

            ime_off()  # フォーカス奪取を避けるため最後に
            return  # 成功
        else:
            time.sleep(1)

    # タイムアウトした場合
    raise MyError('ERROR:' + app_name + 'ウィンドウ取得失敗')


# 画像認識チェック（画像場所,x移動座標,y移動座標,タイムアウト時間）
def img_check(img_name, x_move, y_move, timeout):
    start = time.time()
    while time.time() - start <= 60 * timeout:  # タイムアウト
        try:

            def imread(img_name):  # OpenCVファイルパス日本語変換　参考：https://qiita.com/SKYS/items/cbde3775e2143cad7455
                try:
                    n = np.fromfile(img_name, np.uint8)
                    img = cv2.imdecode(n, cv2.IMREAD_COLOR)
                    return img
                except MyError:
                    raise MyError('ERROR:' + 'ファイルパス日本語変換失敗')

            ag.moveTo(1, 1)  # カーソル移動
            check = ag.locateCenterOnScreen(imread(img_name))  # 画像認識
            if check is None:  # 画像認識されない場合
                check = ag.locateCenterOnScreen(imread(img_name), grayscale=True, confidence=0.95)  # 画像認識精度調整
                if check is None:  # 画像認識されない場合
                    check = ag.locateCenterOnScreen(imread(img_name), grayscale=True, confidence=0.9)  # 画像認識精度調整
                    if check is None:  # 画像認識されない場合
                        raise MyError
        except ag.ImageNotFoundException:
            ag.sleep(0.3)
        else:
            try:
                x, y = check  # 上記画像認識設定で再度確認
                if x and y is None:  # 画像認識されない場合
                    raise MyError
            except MyError:
                ag.sleep(0.3)
            else:
                ag.click(x + x_move, y + y_move)  # センタークリックまたは座標調整
                ime_off()
                time.sleep(0.5)
                break
    if time.time() - start >= 60 * timeout:  # タイムアウトした場合
        raise MyError('ERROR:' + img_name + '画像認識失敗')


# pyautogui日本語対応
def copy_paste(Japanese):
    ime_off()
    pyperclip.copy(Japanese)
    ag.hotkey('ctrl', 'v')


# IME無効化　※Google日本語入力インストール設定必要
def ime_off():
    ag.hotkey('ctrlleft', '0')
    time.sleep(0.5)


# ウィンドウが消えるまで待つ（ウィンドウ名）
def app_wait(app_name):
    time.sleep(1)
    while gw.getWindowsWithTitle(app_name) != []:
        time.sleep(1)
    time.sleep(1)


# アラジンオフィス全ウィンドウ閉じる
def process_close(process_name):
    """指定プロセスを強制終了。存在しない場合のエラーは無視する。"""
    try:
        subprocess.run(['taskkill', '/F', '/T', '/IM', process_name], capture_output=True, text=True, check=False)
    except Exception as e:
        print(f'[process_close] warning: {e}')


def ao_reboot_check():
    app_wait('RemoteApp')
    if gw.getWindowsWithTitle('エラー') != []:
        subprocess.Popen(['taskkill', '/F', '/T', '/IM', 'mstsc.exe'])
        time.sleep(60)
        # アラジンオフィス起動(ショートカット)
        subprocess.Popen(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'link', 'アラジンオフィス販売管理.lnk'), shell=True)


# アラジンオフィスログイン
def ao_login():
    process_close('mstsc.exe')

    def login():
        app_check('ログイン', 1)
        ag.press('tab', presses=3, interval=0.2)
        ag.press('enter')

    # アラジンオフィス起動(ショートカット)
    subprocess.Popen(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'link', 'アラジンオフィス販売管理.lnk'), shell=True)
    ao_reboot_check()

    login()

    # if gw.getWindowsWithTitle("ログイン") != []:
    #     raise MyError("ログイン失敗")


# 日付入力（月,日,月指定,日指定）
def date_input(ms, ds, m, d):
    ime_off()

    date = dt.datetime.now() + relativedelta(months=ms, days=ds, month=m, day=d)  # 参考：https://zenn.dev/wtkn25/articles/python-relativedelta
    ag.typewrite(date.strftime('%Y'))
    ag.press('enter')
    ag.typewrite(date.strftime('%m'))
    ag.press('enter')
    ag.typewrite(date.strftime('%d'))


def _find_csv_by_keyword(base_dir: str, keyword: str) -> str | None:
    """base_dir 直下のCSVから、部分一致 keyword を含む最初のファイル名を返す（なければ None）。"""
    try:
        for f in os.listdir(base_dir):
            if f.lower().endswith('.csv') and keyword in f:
                return os.path.join(base_dir, f)
    except FileNotFoundError:
        pass
    return None


# 明細出力 期間指定
def export_timelimit(window_title, retries: int = 5, delay_sec: float = 5.0):
    """
    明細出力（期間指定）。保存完了（CSV存在）まで最大 retries 回リトライ。
    チェック場所: get_base_path() 直下。存在のみ判定。失敗時はログしてスキップ。
    """

    def _do_once():
        copy_paste(window_title)
        ag.press('enter', presses=2)

        app_check(window_title, 1)
        ag.press('enter')
        date_input(-12, 0, 0, 0)  # 3ヶ月前
        ag.press('enter')
        date_input(0, 0, 0, 0)  # 当月
        ag.press('f1')

        app_check(window_title + '問合せ', 1)
        ag.press('f1')

        # 保存先フォルダ選択
        app_check('出力', 1)
        ag.hotkey('ctrlleft', 'f')
        app_check('名前を付けて保存', 1)

        # 保存ダイアログへの貼り付けは RDP対応版パス
        base_path_env = get_base_path_for_env()
        ag.press('left')
        copy_paste(base_path_env + '\\')
        ag.press('enter')
        app_wait('名前を付けて保存')

        # # 基本パス取得（絶対パス）
        # base_path = os.path.dirname(sys.argv[0])
        # arch = platform.machine()

        # if "ARM" in arch.upper() or "AARCH64" in arch.upper():
        #     # Windowsから見たMacのパスはZ:\Mac\Home\～なので、それをZ:\～にする
        #     rel_path = os.path.relpath(base_path, "C:\\")
        #     adjusted_path = re.sub(r'^Mac\\Home\\', '', rel_path)
        #     copy_paste(os.path.join(r"\\tsclient\Z", adjusted_path) + "\\")
        # else:
        #     copy_paste(base_path + "\\")

        # ag.press("enter")

        app_check('出力', 1)
        ag.press('1')
        ag.press('enter')
        ag.press('2')
        ag.press('enter')
        ag.press('1')
        ag.press('enter')
        ag.press('0')
        ag.press('enter')
        ag.press('1')
        ag.press('enter')
        ag.press('1')
        ag.press('f1')

        app_check('情報', 10)  # 長期間検索用：判定時間10分
        ag.press('enter')

        app_check(window_title + '問合せ', 1)
        ag.press('f12')

        app_check(window_title, 1)
        ag.press('f12')

        img_check(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'image', '001_削除.png'), 0, 0, 1)
        ag.press('tab', presses=2)

    # --- リトライ制御 ---
    keyword = f'{window_title}問合せ'  # ファイル名部分一致想定（例：受注明細問合せ.csv）
    base_dir_chk = get_base_path()

    for attempt in range(1, retries + 1):
        _do_once()
        found = _find_csv_by_keyword(base_dir_chk, keyword)
        if found:
            print(f'[export_timelimit] 保存確認OK: {found}')
            return True
        else:
            if attempt < retries:
                print(f'[export_timelimit] CSV未検出（{keyword}）。{delay_sec:.0f}秒後にリトライ {attempt}/{retries}')
                time.sleep(delay_sec)
            else:
                print(f'[export_timelimit] CSV未検出（{keyword}）。最大リトライ到達のためスキップ')
                return False


# 一覧出力
def export_summary(search_title, window_title, retries: int = 5, delay_sec: float = 5.0):
    """
    一覧出力。保存完了（CSV存在）まで最大 retries 回リトライ。
    チェック場所: get_base_path() 直下。存在のみ判定。失敗時はログしてスキップ。
    """

    def _do_once():
        app_check('アラジンオフィス.NET', 1)
        copy_paste(search_title)
        ag.press('enter', presses=2)

        app_check(window_title, 1)
        ag.press('f1')

        app_check(window_title + '問合せ', 1)
        ag.press('f1')

        # 保存先フォルダ選択
        app_check('出力', 1)
        ag.hotkey('ctrlleft', 'f')
        app_check('名前を付けて保存', 1)

        base_path_env = get_base_path_for_env()
        ag.press('left')
        copy_paste(base_path_env + '\\')
        ag.press('enter')
        app_wait('名前を付けて保存')

        # # 基本パス取得（絶対パス）
        # base_path = os.path.dirname(sys.argv[0])
        # arch = platform.machine()

        # if "ARM" in arch.upper() or "AARCH64" in arch.upper():
        #     # Windowsから見たMacのパスはZ:\Mac\Home\～なので、それをZ:\～にする
        #     rel_path = os.path.relpath(base_path, "C:\\")
        #     adjusted_path = re.sub(r'^Mac\\Home\\', '', rel_path)
        #     copy_paste(os.path.join(r"\\tsclient\Z", adjusted_path) + "\\")
        # else:
        #     copy_paste(base_path + "\\")

        # ag.press("enter")

        app_check('出力', 1)
        ag.press('1')
        ag.press('enter')
        ag.press('2')
        ag.press('enter')
        ag.press('1')
        ag.press('enter')
        ag.press('1')
        ag.press('enter')
        ag.press('1')
        ag.press('f1')

        app_check('情報', 1)
        ag.press('enter')

        app_check(window_title + '問合せ', 1)
        ag.press('f12')

        app_check(window_title, 1)
        ag.press('f12')

        img_check(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'image', '001_削除.png'), 0, 0, 1)
        ag.press('tab', presses=2)

    # --- リトライ制御 ---
    keyword = f'{window_title}問合せ'  # 例：得意先マスタ一覧表問合せ.csv
    base_dir_chk = get_base_path()

    for attempt in range(1, retries + 1):
        _do_once()
        found = _find_csv_by_keyword(base_dir_chk, keyword)
        if found:
            print(f'[export_summary] 保存確認OK: {found}')
            return True
        else:
            if attempt < retries:
                print(f'[export_summary] CSV未検出（{keyword}）。{delay_sec:.0f}秒後にリトライ {attempt}/{retries}')
                time.sleep(delay_sec)
            else:
                print(f'[export_summary] CSV未検出（{keyword}）。最大リトライ到達のためスキップ')
                return False


# # 明細出力 期間指定
# def export_timelimit(window_title):
#     copy_paste(window_title)
#     ag.press('enter', presses=2)

#     app_check(window_title, 1)
#     ag.press('enter')
#     date_input(-3, 0, 0, 0)  # 1ヶ月前
#     ag.press('enter')
#     date_input(0, 0, 0, 0)  # 当月
#     ag.press('f1')

#     app_check(window_title + '問合せ', 1)
#     ag.press('f1')

#     # 保存先フォルダ選択
#     app_check('出力', 1)
#     ag.hotkey('ctrlleft', 'f')
#     app_check('名前を付けて保存', 1)

#     # RDP判定
#     base_path = get_base_path_for_env()

#     ag.press('left')
#     copy_paste(base_path + '\\')

#     ag.press('enter')
#     app_wait('名前を付けて保存')

#     # # 基本パス取得（絶対パス）
#     # base_path = os.path.dirname(sys.argv[0])
#     # arch = platform.machine()

#     # if "ARM" in arch.upper() or "AARCH64" in arch.upper():
#     #     # Windowsから見たMacのパスはZ:\Mac\Home\～なので、それをZ:\～にする
#     #     rel_path = os.path.relpath(base_path, "C:\\")
#     #     adjusted_path = re.sub(r'^Mac\\Home\\', '', rel_path)
#     #     copy_paste(os.path.join(r"\\tsclient\Z", adjusted_path) + "\\")
#     # else:
#     #     copy_paste(base_path + "\\")

#     # ag.press("enter")

#     app_check('出力', 1)
#     ag.press('1')
#     ag.press('enter')
#     ag.press('2')
#     ag.press('enter')
#     ag.press('1')
#     ag.press('enter')
#     ag.press('0')
#     ag.press('enter')
#     ag.press('1')
#     ag.press('enter')
#     ag.press('1')
#     ag.press('f1')

#     app_check('情報', 10)  # 長期間検索用：判定時間10分
#     ag.press('enter')

#     app_check(window_title + '問合せ', 1)
#     ag.press('f12')

#     app_check(window_title, 1)
#     ag.press('f12')

#     img_check(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'image', '001_削除.png'), 0, 0, 1)
#     ag.press('tab', presses=2)


# # 一覧出力
# def export_summary(search_title, window_title):
#     app_check('アラジンオフィス.NET', 1)
#     copy_paste(search_title)
#     ag.press('enter', presses=2)

#     app_check(window_title, 1)
#     ag.press('f1')

#     app_check(window_title + '問合せ', 1)
#     ag.press('f1')

#     # 保存先フォルダ選択
#     app_check('出力', 1)
#     ag.hotkey('ctrlleft', 'f')
#     app_check('名前を付けて保存', 1)

#     # RDP判定
#     base_path = get_base_path_for_env()

#     ag.press('left')
#     copy_paste(base_path + '\\')

#     ag.press('enter')
#     app_wait('名前を付けて保存')

#     # # 基本パス取得（絶対パス）
#     # base_path = os.path.dirname(sys.argv[0])
#     # arch = platform.machine()

#     # if "ARM" in arch.upper() or "AARCH64" in arch.upper():
#     #     # Windowsから見たMacのパスはZ:\Mac\Home\～なので、それをZ:\～にする
#     #     rel_path = os.path.relpath(base_path, "C:\\")
#     #     adjusted_path = re.sub(r'^Mac\\Home\\', '', rel_path)
#     #     copy_paste(os.path.join(r"\\tsclient\Z", adjusted_path) + "\\")
#     # else:
#     #     copy_paste(base_path + "\\")

#     # ag.press("enter")

#     app_check('出力', 1)
#     ag.press('1')
#     ag.press('enter')
#     ag.press('2')
#     ag.press('enter')
#     ag.press('1')
#     ag.press('enter')
#     ag.press('1')
#     ag.press('enter')
#     ag.press('1')
#     ag.press('f1')

#     app_check('情報', 1)
#     ag.press('enter')

#     app_check(window_title + '問合せ', 1)
#     ag.press('f12')

#     app_check(window_title, 1)
#     ag.press('f12')

#     img_check(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'image', '001_削除.png'), 0, 0, 1)
#     ag.press('tab', presses=2)


# アラジン操作
def ao_action():
    app_check('アラジンオフィス.NET', 1)
    img_check(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'image', '000_検索.png'), 0, 0, 1)

    # 明細出力 受注管理
    export_timelimit('受注明細')

    # 明細出力 発注管理
    export_timelimit('発注明細')

    # 明細出力 売上明細（絶対期間）
    today_str = datetime.now().strftime('%Y/%m/%d')  # JST前提の現在日
    export_timelimit_absolute('売上明細', '2024/01/01', today_str)

    # 一覧出力 得意先一覧表
    export_summary('得意先一覧表', '得意先マスタ一覧表')

    # 一覧出力 納品先一覧表
    export_summary('納品先一覧表', '納品先マスタ一覧表')

    # 一覧出力 仕入先一覧表
    export_summary('仕入先一覧表', '仕入先マスタ一覧表')

    # アラジンを閉じる
    process_close('mstsc.exe')


def get_base_path_for_env():
    """
    書き込み基準ディレクトリを返す。
    - RDPセッション中: 実行ディレクトリをそのまま
    - 非RDP       : 実行ディレクトリがローカルドライブなら \\tsclient\\{Drive}\\... に写像
                    実行ディレクトリがUNCならそのまま
    """
    dirpath = os.path.abspath(os.path.dirname(sys.argv[0]))
    is_rdp = os.environ.get('SESSIONNAME', '').upper().startswith('RDP-')

    if is_rdp:
        return dirpath

    drive, tail = os.path.splitdrive(dirpath)
    if drive:  # 例: "C:"
        drive_letter = drive.rstrip(':')
        return os.path.join(r'\\tsclient', drive_letter + tail)
    else:
        # UNC 等（\\server\share\...）はそのまま返す
        return dirpath


def get_base_path() -> str:
    """サーバ側の実行ファイル（スクリプト/exe）と同じディレクトリを返す。"""
    return os.path.abspath(os.path.dirname(sys.argv[0]))


def move_csv_files_from_base(keywords):
    """
    実行ディレクトリ直下のみを探索し、
        - files/data/ に固定名でコピー保存
        - files/csv/ に元名でバックアップ移動（重複時は _YYYYMMDD_HHMMSS）
    """
    base_dir = get_base_path()
    save_dir = os.path.join(base_dir, 'files', 'data')
    backup_dir = os.path.join(base_dir, 'files', 'csv')
    os.makedirs(save_dir, exist_ok=True)
    os.makedirs(backup_dir, exist_ok=True)

    save_name_map = {
        '受注明細問合せ': 'order_upsert.csv',
        '発注明細問合せ': 'p_order_upsert.csv',
        '売上明細問合せ': 'sales_upsert.csv',
        '得意先マスタ一覧表問合せ': 'client_upsert.csv',
        '納品先マスタ一覧表問合せ': 'delivery_upsert.csv',
        '仕入先マスタ一覧表問合せ': 'suppliers_upsert.csv',
    }

    try:
        entries = os.listdir(base_dir)
    except FileNotFoundError:
        print(f'[WARN] 実行ディレクトリが見つかりません: {base_dir}')
        return

    ks = [k.lower() for k in keywords]
    matched = [f for f in entries if f.lower().endswith('.csv') and any(k in f.lower() for k in ks)]

    if not matched:
        print('該当CSVなし。検索パス:', base_dir)
        return

    for filename in matched:
        src = os.path.join(base_dir, filename)

        fixed = next((dst for key, dst in save_name_map.items() if key in filename), None)
        if not fixed:
            print(f'[WARN] 保存名マッピング未定義: {filename}（スキップ）')
            continue

        # 1) data へコピー
        dst_fixed = os.path.join(save_dir, fixed)
        print(f'コピー保存: {src} → {dst_fixed}')
        shutil.copy2(src, dst_fixed)

        # 2) csv へバックアップ移動（重複時は時刻サフィックス）
        dst_backup = os.path.join(backup_dir, filename)
        if os.path.exists(dst_backup):
            ts = datetime.now().strftime('%Y%m%d_%H%M%S')
            name, ext = os.path.splitext(filename)
            dst_backup = os.path.join(backup_dir, f'{name}_{ts}{ext}')
        print(f'バックアップ移動: {src} → {dst_backup}')
        shutil.move(src, dst_backup)


def delete_old_files(folder_path):
    # フォルダの存在を確認する
    folder = Path(folder_path)
    if not folder.exists():
        print(f"Folder '{folder_path}' not found.")
        return  # フォルダが存在しない場合は何もしない

    # 現在の日時を取得し、30日前の日時を計算
    current_time = datetime.now()
    thirty_days_ago = current_time - timedelta(days=30)

    # フォルダ内のファイルをチェックして、30日以上経過したファイルを削除
    for file_name in folder.glob('*'):
        file_modified_time = datetime.fromtimestamp(os.path.getmtime(file_name))
        if file_modified_time < thirty_days_ago:
            os.remove(file_name)


# --------------------------------------------------------------------------------------
# 設定値
# --------------------------------------------------------------------------------------

# スプレッドシート（DB 本体）
spreadsheetId = '1OaegaP4vhLGW-8jRYOaHaoHibJ6p--h82FH-1RA40ag'
dbSheetName = 'DB'

# サービスアカウント認証
BASE_DIR = os.path.dirname(__file__)
SEC_DIR = os.path.join(BASE_DIR, 'files', 'secrets')
credentialsFile = os.path.join(SEC_DIR, 'credentials.json')

# CSV の配置ディレクトリ
filesDir = os.path.join(os.path.dirname(__file__), 'files', 'data')

# 主キー
keyCols = ['受注発注NO', '受注行NO']

# スプレッドシートの許容セル数（Completed の回転判断に使用）
sheetsCellLimit = 10_000_000

# Archive 管理（外部スプレッドシート群）
USER_CLIENT_SECRET = os.path.join(SEC_DIR, 'client_secret.json')
USER_TOKEN_PATH = os.path.join(SEC_DIR, 'token.json')
ARCHIVE_FOLDER_ID = '1LpUF8JuR9ujjc_--YAkXbYn_6HCZb4KI'
SA_CLIENT_EMAIL = 'bd-gsheets-batch-prod@db-gsheets-batch-prod.iam.gserviceaccount.com'
ARCHIVE_INDEX_SHEET_ID = '18veCzW6fxde0SlTk6LGuiJYRFQbbYpeOMJBvIIDktYI'
ARCHIVE_INDEX_TAB_NAME = 'ArchiveIndex'
ARCHIVE_SELECT_POLICY = 'seq'  # or "date"
ARCHIVE_HEADERS = ['seq', 'title', 'fileId', 'createdTime', 'minDate', 'maxDate', 'periodLabel']

# --------------------------------------------------------------------------------------
# 認証・基本ユーティリティ
# --------------------------------------------------------------------------------------


def authorizeGspread(filePath: str) -> gspread.Client:
    """サービスアカウントで gspread クライアントを返す。"""
    scope = [
        'https://www.googleapis.com/auth/spreadsheets',
        'https://www.googleapis.com/auth/drive',
    ]
    # gspread.service_account は Client を直接返すので、そのまま返す
    return gspread.service_account(filename=filePath, scopes=scope)


def authorize_user_drive(client_secret_path: str, token_path: str, scopes=None):
    """ユーザー OAuth（Installed App）で Drive API サービスを返す。"""
    if scopes is None:
        scopes = ['https://www.googleapis.com/auth/drive']

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, scopes)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_path, scopes)
            creds = flow.run_local_server(port=0)
        with open(token_path, 'w', encoding='utf-8') as f:
            f.write(creds.to_json())

    return build('drive', 'v3', credentials=creds)


def getSheetDf(sheet) -> pd.DataFrame:
    """ワークシート全体を DataFrame 化（ヘッダ 1 行）。空なら空 DF。"""
    vals = sheet.get_all_values()
    if not vals or len(vals) < 2:
        return pd.DataFrame(columns=vals[0] if vals else [])
    return pd.DataFrame(vals[1:], columns=vals[0])


# --------------------------------------------------------------------------------------
# Completed シートの確保／外部アーカイブの選定と回転
# --------------------------------------------------------------------------------------


def ensure_completed_sheet(spreadsheet, min_rows=2, min_cols=10):
    """ブック内に 'Completed' を必ず用意（既存なら軽くリサイズ）。"""
    try:
        ws = spreadsheet.worksheet('Completed')
        try:
            if ws.row_count < min_rows or ws.col_count < min_cols:
                ws.resize(rows=max(ws.row_count, min_rows), cols=max(ws.col_count, min_cols))
        except APIError as e:
            print(f'[ensure_completed_sheet] resize warning: {e}')
        return ws
    except WorksheetNotFound:
        pass

    # 既存の最初のシートをリネーム出来れば採用、不可なら新規作成
    try:
        wss = spreadsheet.worksheets()
    except Exception as e:
        print(f'[ensure_completed_sheet] list worksheets failed: {e}')
        wss = []

    if wss:
        ws0 = wss[0]
        try:
            ws0.update_title('Completed')
            try:
                if ws0.row_count < min_rows or ws0.col_count < min_cols:
                    ws0.resize(rows=max(ws0.row_count, min_rows), cols=max(ws0.col_count, min_cols))
            except APIError as e:
                print(f'[ensure_completed_sheet] resize warning: {e}')
            return ws0
        except APIError as e:
            print(f'[ensure_completed_sheet] rename failed, fallback to add new: {e}')

    return spreadsheet.add_worksheet(title='Completed', rows=min_rows, cols=min_cols)


def _open_archive_index_ws(client):
    ss = client.open_by_key(ARCHIVE_INDEX_SHEET_ID)
    try:
        ws = ss.worksheet(ARCHIVE_INDEX_TAB_NAME)
    except WorksheetNotFound:
        ws = ss.add_worksheet(title=ARCHIVE_INDEX_TAB_NAME, rows=2, cols=len(ARCHIVE_HEADERS))
        ws.update('A1:G1', [ARCHIVE_HEADERS])
    return ws


def _read_archive_index_df(client):
    ws = _open_archive_index_ws(client)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        return pd.DataFrame(columns=ARCHIVE_HEADERS), ws
    header = vals[0][: len(ARCHIVE_HEADERS)]
    body = [row[: len(ARCHIVE_HEADERS)] for row in vals[1:]]
    df = pd.DataFrame(body, columns=header)
    if 'seq' in df.columns:
        df['seq'] = pd.to_numeric(df['seq'], errors='coerce').fillna(0).astype(int)
    for c in ('minDate', 'maxDate', 'createdTime'):
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors='coerce')
    return df, ws


def _pick_latest_archive_row(df: pd.DataFrame, policy: str):
    if df.empty:
        return None
    cols = set(df.columns)
    if policy == 'date':
        if 'maxDate' in cols:
            # Build sort keys explicitly to avoid starred conditional expressions
            sort_keys = ['maxDate']
            if 'seq' in cols:
                sort_keys.append('seq')
            return df.sort_values(sort_keys).tail(1).iloc[0]
        if 'createdTime' in cols:
            return df.sort_values(['createdTime']).tail(1).iloc[0]
        return df.tail(1).iloc[0]
    if 'seq' in cols:
        return df.sort_values(['seq']).tail(1).iloc[0]
    if 'createdTime' in cols:
        return df.sort_values(['createdTime']).tail(1).iloc[0]
    return df.tail(1).iloc[0]


def _find_index_rownum_by_file_id(ws, file_id: str) -> int | None:
    vals = ws.get_all_values()
    if not vals:
        return None
    header = vals[0]
    try:
        col_idx = header.index('fileId')
    except ValueError:
        return None
    for i, row in enumerate(vals[1:], start=2):
        if len(row) > col_idx and row[col_idx] == file_id:
            return i
    return None


def _update_index_metadata(ws, rownum: int, *, min_d, max_d, period_label: str):
    min_str = '' if pd.isna(min_d) or min_d is None else pd.to_datetime(min_d).strftime('%Y-%m-%d')
    max_str = '' if pd.isna(max_d) or max_d is None else pd.to_datetime(max_d).strftime('%Y-%m-%d')
    ws.update(range_name=f'E{rownum}:G{rownum}', values=[[min_str, max_str, period_label]])


def create_spreadsheet_on_mydrive(drive_service, folder_id: str, title: str) -> str:
    metadata = {'name': title, 'mimeType': 'application/vnd.google-apps.spreadsheet'}
    if folder_id:
        metadata['parents'] = [folder_id]
    created = drive_service.files().create(body=metadata, fields='id').execute()
    return created['id']


def grant_editor_to_sa(drive_service, file_id: str, sa_email: str):
    perm = {'type': 'user', 'role': 'writer', 'emailAddress': sa_email}
    drive_service.permissions().create(fileId=file_id, body=perm, sendNotificationEmail=False, fields='id').execute()


def ensure_current_archive_spreadsheet(client, comp_new_df):
    """ArchiveIndex を参照し、現在の Completed 保存先ブックを返す（無ければ新規作成）。"""
    df, ws = _read_archive_index_df(client)
    row = _pick_latest_archive_row(df, ARCHIVE_SELECT_POLICY)

    # 期間情報の推定
    if '受注日' in comp_new_df.columns:
        s = pd.to_datetime(comp_new_df['受注日'], errors='coerce')
        min_d = s.min() if s.notna().any() else None
        max_d = s.max() if s.notna().any() else None
    else:
        min_d = max_d = None
    period = pd.to_datetime(min_d).strftime('%Y%m') if min_d is not None else datetime.now().strftime('%Y%m')

    if row is not None:
        file_id = row['fileId']
        ss = client.open_by_key(file_id)
        ensure_completed_sheet(ss)
        if min_d is not None or max_d is not None:
            rn = _find_index_rownum_by_file_id(ws, file_id)
            if rn:
                _update_index_metadata(ws, rn, min_d=min_d, max_d=max_d, period_label=period)
        return ss

    # 初回: 新規作成 → SA 付与 → Index 追記
    seq = 1
    title = f'Completed_Archive_{seq:04d}_{period}'
    user_drive = authorize_user_drive(USER_CLIENT_SECRET, USER_TOKEN_PATH)
    new_id = create_spreadsheet_on_mydrive(user_drive, ARCHIVE_FOLDER_ID, title)
    grant_editor_to_sa(user_drive, new_id, SA_CLIENT_EMAIL)
    created_time = pd.to_datetime(user_drive.files().get(fileId=new_id, fields='createdTime').execute()['createdTime'])
    ws.append_row([str(seq), title, new_id, created_time.isoformat(), '' if min_d is None else pd.to_datetime(min_d).strftime('%Y-%m-%d'), '' if max_d is None else pd.to_datetime(max_d).strftime('%Y-%m-%d'), period])
    ss = client.open_by_key(new_id)
    ensure_completed_sheet(ss)
    return ss


def maybe_rotate_archive_spreadsheet(client, archive_ss, rows_to_add: int, comp_new_df: pd.DataFrame):
    """Completed の容量しきい超過で外部アーカイブを回転（新規作成）する。"""
    try:
        comp_ws = archive_ss.worksheet('Completed')
    except WorksheetNotFound:
        comp_ws = archive_ss.add_worksheet(title='Completed', rows=2, cols=10)

    cols = max(1, comp_ws.col_count)
    threshold = sheetsCellLimit // cols
    try:
        current_rows = max(0, len(comp_ws.get_all_values()) - 1)
    except Exception:
        current_rows = max(0, comp_ws.row_count - 1)

    # 期間情報
    if '受注日' in comp_new_df.columns:
        s = pd.to_datetime(comp_new_df['受注日'], errors='coerce')
        min_d = s.min() if s.notna().any() else None
        max_d = s.max() if s.notna().any() else None
    else:
        min_d = max_d = None
    period = pd.to_datetime(min_d).strftime('%Y%m') if min_d is not None else datetime.now().strftime('%Y%m')

    df, ws = _read_archive_index_df(client)

    if current_rows + rows_to_add <= threshold:
        file_id = archive_ss.id
        rn = _find_index_rownum_by_file_id(ws, file_id)
        if rn and (min_d is not None or max_d is not None):
            _update_index_metadata(ws, rn, min_d=min_d, max_d=max_d, period_label=period)
        return archive_ss

    # 新規作成
    seq = (int(df['seq'].max()) + 1) if not df.empty else 1
    title = f'Completed_Archive_{seq:04d}_{period}'
    user_drive = authorize_user_drive(USER_CLIENT_SECRET, USER_TOKEN_PATH)
    new_id = create_spreadsheet_on_mydrive(user_drive, ARCHIVE_FOLDER_ID, title)
    grant_editor_to_sa(user_drive, new_id, SA_CLIENT_EMAIL)
    created_time = pd.to_datetime(user_drive.files().get(fileId=new_id, fields='createdTime').execute()['createdTime'])
    ws.append_row([str(seq), title, new_id, created_time.isoformat(), '' if min_d is None else pd.to_datetime(min_d).strftime('%Y-%m-%d'), '' if max_d is None else pd.to_datetime(max_d).strftime('%Y-%m-%d'), period])
    return client.open_by_key(new_id)


def ensure_archive_and_maybe_rotate(client, archive_state: dict, rows_to_add: int, date_src_df: pd.DataFrame):
    """現在のアーカイブ先を確定し、必要なら回転する。戻り値は書き込み先ブック。"""
    if archive_state.get('ss') is None:
        archive_state['ss'] = ensure_current_archive_spreadsheet(client, date_src_df)
        archive_state['rotated'] = False
    before_id = archive_state['ss'].id
    ss2 = maybe_rotate_archive_spreadsheet(client, archive_state['ss'], rows_to_add, date_src_df)
    archive_state['ss'] = ss2
    archive_state['rotated'] = ss2.id != before_id
    ensure_completed_sheet(archive_state['ss'])  # 念のため
    return archive_state['ss']


# --------------------------------------------------------------------------------------
# Completed_Archive を DB スキーマに正規化するユーティリティ
# --------------------------------------------------------------------------------------


def _normalize_completed_ws_to_db_schema(ws, db_cols: list[str]) -> None:
    """
    指定ワークシート(Completed)の列集合・順序を DB の列(db_cols)に完全一致させる。
    - Archive にあって DB にない列は削除
    - DB にあって Archive にない列は空で追加
    - 最終的に列順も db_cols と同一に並べ替え
    """
    df = getSheetDf(ws)

    # シートが完全空/ヘッダなしの場合は空DFにDB列だけを持たせて書き戻す
    if df.shape[1] == 0:
        empty_df = pd.DataFrame(columns=db_cols)
        clearAndWriteDf(ws, empty_df)
        return

    # 追加（DBにあってArchiveにない列）
    for c in db_cols:
        if c not in df.columns:
            df[c] = ''

    # 削除（ArchiveにあってDBにない列）
    extra_cols = [c for c in df.columns if c not in db_cols]
    if extra_cols:
        df = df.drop(columns=extra_cols, errors='ignore')

    # 並べ替え（DBと同じ順序）
    df = df.reindex(columns=db_cols, fill_value='')

    # 書き戻し
    clearAndWriteDf(ws, df)


def normalize_all_archives_to_db_schema(client: gspread.Client, db_cols: list[str]) -> None:
    """
    ArchiveIndex に登録されている全アーカイブブックの Completed を走査し、
    それぞれの列集合・順序を DB の列(db_cols)に完全一致させる。
    """
    # ArchiveIndex を取得
    df_idx, _ = _read_archive_index_df(client)
    if df_idx.empty or 'fileId' not in df_idx.columns:
        print('[normalize_all_archives_to_db_schema] ArchiveIndex が空、または fileId 列なし。処理をスキップします。')
        return

    file_ids = df_idx['fileId'].dropna().astype(str).unique().tolist()
    for fid in file_ids:
        try:
            ss = client.open_by_key(fid)
        except Exception as e:
            print(f'[normalize_all_archives_to_db_schema] open_by_key 失敗: {fid} ({e})')
            continue

        # Completed を必ず用意
        try:
            ws = ss.worksheet('Completed')
        except WorksheetNotFound:
            ws = ss.add_worksheet(title='Completed', rows=2, cols=max(10, len(db_cols)))

        # 正規化実行
        try:
            _normalize_completed_ws_to_db_schema(ws, db_cols)
            print(f'[normalize_all_archives_to_db_schema] 正規化完了: {fid}')
        except Exception as e:
            print(f'[normalize_all_archives_to_db_schema] 正規化失敗: {fid} ({e})')


# --------------------------------------------------------------------------------------
# Conflicts 管理
# --------------------------------------------------------------------------------------


def ensure_named_sheet(spreadsheet, title, min_rows=2, min_cols=10):
    try:
        ws = spreadsheet.worksheet(title)
        try:
            if ws.row_count < min_rows or ws.col_count < min_cols:
                ws.resize(rows=max(ws.row_count, min_rows), cols=max(ws.col_count, min_cols))
        except APIError as e:
            print(f'[ensure_named_sheet] resize warning: {e}')
        return ws
    except WorksheetNotFound:
        return spreadsheet.add_worksheet(title=title, rows=min_rows, cols=min_cols)


def ensure_conflicts_sheet(spreadsheet, min_rows=2, min_cols=10):
    return ensure_named_sheet(spreadsheet, 'Conflicts', min_rows=min_rows, min_cols=min_cols)


def _append_reason(df, *, reason_code, reason_detail='', source=''):
    out = df.copy()
    out['reason_code'] = reason_code
    out['reason_detail'] = reason_detail
    out['source'] = source
    out['timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    return out


def _detect_csv_dup_keys(df, key_cols, source_label):
    """CSV 内の主キー重複（同一 受注発注NO×受注行NO）を検出し、重複のみ抽出。"""
    if df.empty:
        return df.copy(), pd.DataFrame(columns=df.columns)
    dup_mask = df.duplicated(subset=key_cols, keep=False)
    dup_df = df.loc[dup_mask].copy()
    keep_df = df.loc[~dup_mask].copy()
    if not dup_df.empty:
        dup_df = _append_reason(
            dup_df,
            reason_code='DUP_KEY_IN_CSV',
            reason_detail=f'キー重複: {key_cols}',
            source=source_label,
        )
    return keep_df, dup_df


def _unique_left_match(left_df, right_df, *, on=None, left_on=None, right_on=None, how='left', indicator=True, match_label=''):
    """左結合後、左キー単位で複数行に増えたものを“曖昧”として分離。
    戻り値: (safe_df, ambiguous_df, left_only_df)
    """
    m = pd.merge(left_df, right_df, on=on, left_on=left_on, right_on=right_on, how=how, indicator=indicator)
    m = m.loc[:, ~m.columns.duplicated()].copy()
    left_only = m[m['_merge'] == 'left_only'].copy() if '_merge' in m.columns else m.iloc[0:0].copy()
    both = m[m['_merge'] == 'both'].copy() if '_merge' in m.columns else m.copy()

    key_cols = [c for c in ['受注発注NO', '受注行NO'] if c in left_df.columns]
    if not key_cols:
        safe = both.copy()
        amb = both.iloc[0:0].copy()
    else:
        amb_mask = both.duplicated(subset=key_cols, keep=False)
        amb = both.loc[amb_mask].copy()
        safe = both.loc[~amb_mask].copy()

    if match_label:
        if not safe.empty:
            safe['matchMethod'] = match_label
        if not amb.empty:
            amb['matchMethod'] = match_label
    return safe, amb, left_only


# --------------------------------------------------------------------------------------
# CSV→DB のメイン処理
# --------------------------------------------------------------------------------------
# --- Acceptance helpers (ADD) -----------------------------------------------
def _nz(v: any) -> str:
    """None/NaNを空文字に、他はstrでトリムして返す。"""
    if pd.isna(v):
        return ''
    return str(v).strip()


def _zero_pad3(s: str) -> str:
    """行NOの比較用（'1' -> '001' など）。数値だけ3桁ゼロ埋め、その他はそのまま。"""
    t = _nz(s).strip()
    if not t:
        return ''
    if t.isdigit():
        return f'{int(t):03d}'
    return t


def build_acceptance_context(
    activeDf: pd.DataFrame,
    allComp: pd.DataFrame,
    final_un: pd.DataFrame,
    *,
    order_csv: pd.DataFrame | None = None,
    keyCols=('受注発注NO', '受注行NO'),
):
    """
    受け入れ判定に必要な索引をまとめて構築。
    追加:
      - keys_in_order_csv: order_upsert.csv（受注側）に存在するキー集合
        → 「受注側に無い行は受け入れない」判定に使用
    戻り値: dict(headers_in_db, line001_in_db_or_comp, line001_in_batch, activeDf_index, keys_in_order_csv)
    """
    k0, k1 = keyCols

    # DBに存在するヘッダ（受注発注NO）
    headers_in_db = set(activeDf[k0].dropna().astype(str).str.strip()) if k0 in activeDf.columns else set()

    # DB/Completed に行NO=001 が存在するヘッダ
    def _has_001(df):
        if df is None or df.empty:
            return set()
        if k0 not in df.columns or k1 not in df.columns:
            return set()
        t = df[[k0, k1]].copy()
        t[k0] = t[k0].fillna('').astype(str).str.strip()
        t[k1] = t[k1].fillna('').astype(str).str.strip().map(_zero_pad3)
        return set(t.loc[t[k1] == '001', k0])

    line001_in_db = _has_001(activeDf)
    line001_in_comp = _has_001(allComp)
    line001_in_db_or_comp = line001_in_db | line001_in_comp

    # 今回バッチ（NO_DB_MATCH候補群）内で 001 を含むヘッダ
    line001_in_batch = set()
    if final_un is not None and not final_un.empty and k0 in final_un.columns and k1 in final_un.columns:
        t = final_un[[k0, k1]].copy()
        t[k0] = t[k0].fillna('').astype(str).str.strip()
        t[k1] = t[k1].fillna('').astype(str).str.strip().map(_zero_pad3)
        line001_in_batch = set(t.loc[t[k1] == '001', k0])

    # DBのキー→行参照（冪等確認用）
    if k0 in activeDf.columns and k1 in activeDf.columns:
        _idx = activeDf[[k0, k1]].copy()
        _idx[k0] = _idx[k0].fillna('').astype(str).str.strip()
        _idx[k1] = _idx[k1].fillna('').astype(str).str.strip().map(_zero_pad3)
        activeDf_index = activeDf.copy()
        activeDf_index[k0] = _idx[k0].values
        activeDf_index[k1] = _idx[k1].values
        activeDf_index = activeDf_index.set_index([k0, k1])
    else:
        activeDf_index = pd.DataFrame().set_index([k0, k1])

    # 受注CSV（order_upsert.csv）に存在するキー集合
    keys_in_order_csv: set[tuple[str, str]] = set()
    if order_csv is not None and not order_csv.empty and k0 in order_csv.columns and k1 in order_csv.columns:
        t = order_csv[[k0, k1]].copy()
        t[k0] = t[k0].fillna('').astype(str).str.strip()
        t[k1] = t[k1].fillna('').astype(str).str.strip().map(_zero_pad3)
        # 空キーは除外
        t = t[(t[k0] != '') & (t[k1] != '')]
        keys_in_order_csv = set(map(tuple, t.values.tolist()))

    return {
        'headers_in_db': headers_in_db,
        'line001_in_db_or_comp': line001_in_db_or_comp,
        'line001_in_batch': line001_in_batch,
        'activeDf_index': activeDf_index,
        'keys_in_order_csv': keys_in_order_csv,
    }


def _is_idempotent_same_as_db(row: pd.Series, ctx: dict, p_cols: list[str], keyCols=('受注発注NO', '受注行NO')) -> bool:
    """
    R3: 冪等判定。DBに同キーがあり、かつ「DB側に存在する同名カラム」の値が全て一致なら True。
    （hash列をDBに持たず、'実質比較'で等価判定する最小実装）
    """
    k0, k1 = keyCols
    idx = ctx['activeDf_index']
    if idx.empty:
        return False
    key = (_nz(row.get(k0)), _nz(_zero_pad3(row.get(k1))))
    try:
        db_row = idx.loc[key]
    except KeyError:
        return False

    # 比較対象は「候補(p_cols) ∩ DB列」
    common = [c for c in p_cols if c in db_row.index]
    if not common:
        return False

    # 文字列正規化で比較（数値/ゼロ埋め/空文字の差を吸収）
    for c in common:
        if _nz(row.get(c)) != _nz(db_row.get(c)):
            return False
    return True


def _evaluate_accept_row(row: pd.Series, ctx: dict, keyCols=('受注発注NO', '受注行NO')) -> tuple[bool, str, str]:
    """
    R1→R2→R4 を評価。戻り値:
    (accept:bool, reason_code:str, reason_detail:str)
    R3（冪等）は呼び出し側で別関数で付与する。
    """
    k0, k1 = keyCols
    order_id = _nz(row.get(k0))
    line_no = _zero_pad3(row.get(k1))
    key_disp = _fmt_key(order_id, line_no)

    # R1 ヘッダ分散禁止
    if order_id in ctx['headers_in_db']:
        return (False, 'HEADER_SPLIT_EXISTS_IN_DB', f'DBに同一ヘッダ存在のため禁止: {key_disp}')

    # R2 行001必須
    has001_dbcomp = order_id in ctx['line001_in_db_or_comp']
    has001_batch = order_id in ctx['line001_in_batch']
    if not (has001_dbcomp or has001_batch):
        src = []
        if not has001_dbcomp:
            src.append('DB/Completedに001なし')
        if not has001_batch:
            src.append('今回候補に001なし')
        return (False, 'NO_LINE001_PRESENT', f'001行が見つからないため受入不可: {key_disp}（{"・".join(src)}）')

    # R4 最小要件（キー欠損）
    if order_id == '':
        return (False, 'MISSING_ORDER_ID', f'受注発注NOが空のため受入不可: {key_disp}')
    if line_no == '' or not line_no.isdigit():
        return (False, 'MISSING_LINE_ID', f'行NOが不正のため受入不可: {key_disp}')

    return (True, '', '')


# --- Reason detail helpers (ADD) --------------------------------------------
def _fmt_key(order_id: str, line_no: str) -> str:
    """キー表記の統一（例: 00053899-003）"""
    return f'{_nz(order_id)}-{_zero_pad3(line_no)}'


def _df_keys(df: pd.DataFrame, keyCols) -> pd.DataFrame:
    """df から主キー2列だけを返す（欠け列は空で補完）。行NOは3桁ゼロ埋めで正規化。"""
    k0, k1 = keyCols
    if df is None or df.empty:
        return pd.DataFrame(columns=[k0, k1])

    out = df.copy()

    # 列が無ければ作る
    if k0 not in out.columns:
        out[k0] = ''
    if k1 not in out.columns:
        out[k1] = ''

    # 正規化
    out[k0] = out[k0].fillna('').astype(str).str.strip()
    out[k1] = out[k1].fillna('').astype(str).str.strip().map(_zero_pad3)

    return out[[k0, k1]].copy()


def _make_key_tuple_series(df_keys: pd.DataFrame, keyCols) -> list[tuple[str, str]]:
    """(k0,k1) のタプルSeriesを作る（高速な isin 用）。"""
    k0, k1 = keyCols
    s0 = df_keys[k0].fillna('').astype(str)
    s1 = df_keys[k1].fillna('').astype(str)
    return list(zip(s0, s1, strict=False))


def _filter_out_keys(df: pd.DataFrame, keys_df: pd.DataFrame, keyCols) -> pd.DataFrame:
    if df is None or df.empty or keys_df is None or keys_df.empty:
        return df
    k0, k1 = keyCols
    left = _df_keys(df, keyCols)
    keys = _df_keys(keys_df, keyCols)
    left_t = pd.Series(_make_key_tuple_series(left, keyCols), index=df.index)
    keys_t = set(_make_key_tuple_series(keys, keyCols))
    keep_mask = ~left_t.isin(keys_t)
    return df.loc[keep_mask].copy()


def _inner_keys(dfA: pd.DataFrame, dfB: pd.DataFrame, keyCols) -> pd.DataFrame:
    """dfA ∩ dfB のキー集合を返す。"""
    if dfA.empty or dfB.empty:
        return pd.DataFrame(columns=list(keyCols))
    k0, k1 = keyCols
    a = _df_keys(dfA, keyCols).drop_duplicates()
    b = _df_keys(dfB, keyCols).drop_duplicates()
    return a.merge(b, on=[k0, k1], how='inner')


def _ensure_cols(df: pd.DataFrame, cols: list[str], fill_value: str = '') -> pd.DataFrame:
    """df に cols が無ければ追加（空埋め）。"""
    if df is None:
        return pd.DataFrame(columns=list(dict.fromkeys(list(cols))))

    out = df.copy()
    for c in cols:
        if c not in out.columns:
            out[c] = fill_value
    return out


def _dedupe_key_df(df: pd.DataFrame, keyCols) -> pd.DataFrame:
    """キーDFを正規化（列保証 + 重複排除 + 欠損除外）。行NOは3桁ゼロ埋め。"""
    k0, k1 = keyCols
    if df is None or df.empty:
        return pd.DataFrame(columns=[k0, k1])
    out = df.copy()
    out = _ensure_cols(out, [k0, k1], fill_value='')
    out[k0] = out[k0].fillna('').astype(str).str.strip()
    out[k1] = out[k1].fillna('').astype(str).str.strip().map(_zero_pad3)
    out = out[(out[k0] != '') & (out[k1] != '')]
    return out[[k0, k1]].drop_duplicates().reset_index(drop=True)


def _left_only_keys(a_keys: pd.DataFrame, b_keys: pd.DataFrame, keyCols) -> pd.DataFrame:
    """a_keys から b_keys に存在しないキーだけ返す（正規化済み前提）"""
    k0, k1 = keyCols
    if a_keys is None or a_keys.empty:
        return pd.DataFrame(columns=[k0, k1])
    if b_keys is None or b_keys.empty:
        return a_keys[[k0, k1]].copy()

    t = a_keys.merge(b_keys.assign(_in=True), on=[k0, k1], how='left')
    out = t[t['_in'].isna()][[k0, k1]].copy()
    return out.reset_index(drop=True)


def run_import_without_log(client, sh, dbWs, filesDir, keyCols=('受注発注NO', '受注行NO')):
    """CSV 群を取り込み、Completed を外部アーカイブへ、Active を DB に安全更新する。"""
    # --- 既存 DB の読み込みと Completed/Active 分割 ---
    dbDf = getSheetDf(dbWs)
    dbDf = add_missing_columns(dbDf, list(keyCols))

    # --- DBシートの “*_csv” 残骸を除去してから origCols を確定
    #   既存DBの値を尊重したいので prefer="base" 推奨 ---
    dbDf = collapse_csv_suffix(dbDf, prefer='base')
    dbDf = dedupeCols(dbDf)

    origCols = list(dbDf.columns)

    # --- 全 Completed_Archive を DB スキーマ（origCols）に正規化 ---
    normalize_all_archives_to_db_schema(client, origCols)

    compNew, activeDf = splitCompletedActive(dbDf)

    # --- アーカイブ候補の確定（暫定 Completed を元に） ---
    archive_state = {'ss': None, 'rotated': False}
    date_src_A = compNew if not compNew.empty else dbDf
    archive_sh = ensure_archive_and_maybe_rotate(client, archive_state, rows_to_add=len(compNew), date_src_df=date_src_A)

    # Active 側のキー担保
    activeDf = add_missing_columns(activeDf, list(keyCols))

    # --- Completed（外部）を読み込み（この時点では“書き込みしない”） ---
    compWs = ensure_completed_sheet(
        archive_sh,
        min_rows=max(2, len(compNew) + 5),
        min_cols=max(10, len(dbDf.columns) + 5),
    )
    compOld = add_missing_columns(getSheetDf(compWs), list(keyCols))

    # acceptance_context 等で参照するため、メモリ上の allComp は作る（ただし外部へはまだ書かない）
    expCols = origCols + [c for c in compOld.columns if c not in origCols]
    allComp = pd.concat([compOld, compNew], ignore_index=True)
    allComp = allComp.reindex(columns=expCols, fill_value='')
    allComp = allComp.drop_duplicates(subset=list(keyCols), keep='last')
    allComp = allComp.apply(lambda s: s.map(cleanValue))

    # --- CSV 読み込み（列の接頭語付与） ---
    order_csv = loadCsvWithPrefix(os.path.join(filesDir, 'order_upsert.csv'), '受注', excludeCols=list(keyCols))
    porder_csv = loadCsvWithPrefix(os.path.join(filesDir, 'p_order_upsert.csv'), '発注', excludeCols=['発注NO', '発注行NO'])
    client_csv = loadCsvWithPrefix(os.path.join(filesDir, 'client_upsert.csv'), '得意先', excludeCols=['得意先'])
    delivery_csv = loadCsvWithPrefix(os.path.join(filesDir, 'delivery_upsert.csv'), '納品先', excludeCols=['納品先'])
    supplier_csv = loadCsvWithPrefix(os.path.join(filesDir, 'suppliers_upsert.csv'), '仕入先', excludeCols=['仕入先'])

    # キー正規化・必須列保証
    porder_csv = rename_if_exists(porder_csv, {'発注NO': keyCols[0], '発注行NO': keyCols[1]})
    porder_csv = ensure_columns(porder_csv, list(keyCols) + ['発注商品'])
    order_csv = ensure_columns(order_csv, list(keyCols) + ['受注商品'])
    activeDf = ensure_columns(activeDf, list(keyCols) + ['受注商品', '発注商品'])

    # 念のため（列名統一）
    porder_csv = rename_if_exists(porder_csv, {'発注NO': '受注発注NO', '発注行NO': '受注行NO'})

    # --- CSV 内主キー重複を先に Conflicts へ分離 ---
    conflicts = []
    extra_safe_parts = []
    p_cols = list(porder_csv.columns)
    o_cols = list(order_csv.columns)

    porder_csv, dupP = _detect_csv_dup_keys(porder_csv, list(keyCols), source_label='p_order_upsert.csv')
    order_csv, dupO = _detect_csv_dup_keys(order_csv, list(keyCols), source_label='order_upsert.csv')
    if not dupP.empty:
        conflicts.append(dupP)
    if not dupO.empty:
        conflicts.append(dupO)

    # --- 突合（行NO → 商品 → DB 行NO → DB 商品）---
    # (1) 行NO 同士
    m1_safe, m1_amb, m1_left = _unique_left_match(porder_csv, order_csv, on=list(keyCols), how='left', indicator=True, match_label='行NOマッチ')
    if not m1_amb.empty:
        conflicts.append(_append_reason(m1_amb[p_cols], reason_code='AMBIGUOUS_ROW_MATCH', reason_detail='発注×受注(行NO)が1対多', source='p_order_upsert.csv'))

    # (2) 商品名で補完
    order_for_m2 = drop_if_exists(order_csv, [keyCols[1]])
    m2_safe, m2_amb, m2_left = _unique_left_match(m1_left[p_cols], order_for_m2, left_on=[keyCols[0], '発注商品'], right_on=[keyCols[0], '受注商品'], how='left', indicator=True, match_label='商品マッチ')
    if not m2_amb.empty:
        conflicts.append(_append_reason(m2_amb[p_cols], reason_code='AMBIGUOUS_PRODUCT_MATCH', reason_detail='発注×受注(商品)が1対多', source='p_order_upsert.csv'))

    # (3) DB 行NO で再照合
    db_row_tmp = pd.merge(
        m2_left[p_cols],
        activeDf[list(keyCols)],  # 必要最小限に絞る
        on=list(keyCols),
        how='inner',
        suffixes=('_csv', ''),
    )
    if not db_row_tmp.empty:
        dup_mask = db_row_tmp.duplicated(subset=list(keyCols), keep=False)
        db_row_amb = db_row_tmp.loc[dup_mask].copy()
        db_row_safe = db_row_tmp.loc[~dup_mask].copy()
    else:
        db_row_amb = db_row_tmp.iloc[0:0].copy()
        db_row_safe = db_row_tmp.copy()
    if not db_row_amb.empty:
        conflicts.append(_append_reason(db_row_amb[p_cols], reason_code='DB_AMBIGUOUS_ROW_MATCH', reason_detail='Active と(行NO)で多重一致', source='p_order_upsert.csv'))
    db_row_safe['matchMethod'] = 'DB行NOマッチ'

    # (4) DB 商品で再照合
    remaining_after_row = m2_left[p_cols].merge(db_row_safe[list(keyCols)], on=list(keyCols), how='left', indicator=True)
    still_left = remaining_after_row[remaining_after_row['_merge'] == 'left_only'][p_cols].copy()

    db_for_prod = activeDf.drop(columns=[keyCols[1]], errors='ignore')
    _right_needed = [keyCols[0], '受注商品']  # 必要列だけ残す
    _right_needed = [c for c in _right_needed if c in db_for_prod.columns]
    db_prod_tmp = pd.merge(
        still_left,
        db_for_prod[_right_needed],  # 最小限
        left_on=[keyCols[0], '発注商品'],
        right_on=[keyCols[0], '受注商品'],
        how='inner',
        suffixes=('_csv', ''),
    )
    if not db_prod_tmp.empty:
        dup_mask = db_prod_tmp.duplicated(subset=list(keyCols), keep=False)
        db_prod_amb = db_prod_tmp.loc[dup_mask].copy()
        db_prod_safe = db_prod_tmp.loc[~dup_mask].copy()
    else:
        db_prod_amb = db_prod_tmp.iloc[0:0].copy()
        db_prod_safe = db_prod_tmp.copy()
    if not db_prod_amb.empty:
        conflicts.append(_append_reason(db_prod_amb[p_cols], reason_code='DB_AMBIGUOUS_PRODUCT_MATCH', reason_detail='Active と(商品)で多重一致', source='p_order_upsert.csv'))
    db_prod_safe['matchMethod'] = 'DB商品マッチ'

    # (5) なお未確定は Conflicts 扱い（ただし“受け入れ可”は DB へ）
    matched_keys = pd.concat(
        [
            m1_safe[list(keyCols)],
            m2_safe[list(keyCols)],
            db_row_safe[list(keyCols)],
            db_prod_safe[list(keyCols)],
        ],
        ignore_index=True,
    ).drop_duplicates()

    final_un = porder_csv.merge(matched_keys, on=list(keyCols), how='left', indicator=True)
    final_un = final_un[final_un['_merge'] == 'left_only'][p_cols].copy()

    # 未確定行が無ければ A/B/C 判定は不要
    if final_un.empty:
        pass
    else:
        # 受注側欠落の扱い（A/B/C） ---------------------------------
        k0, k1 = keyCols

        # 発注側：ヘッダごとの「001の有無」（A: 002始まり排除）
        p_keys = _df_keys(porder_csv, keyCols)
        p_has001 = set(p_keys.loc[p_keys[k1] == '001', k0])

        # 受注側：ヘッダの存在と、(ヘッダ,行NO) の存在（B/C）
        o_keys = _df_keys(order_csv, keyCols)
        o_headers = set(o_keys[k0].unique())
        o_keyset = set(zip(o_keys[k0].tolist(), o_keys[k1].tolist(), strict=False))

        # final_un をキー正規化
        fu_keys = _df_keys(final_un, keyCols)
        fu_hdr = fu_keys[k0]
        fu_ln = fu_keys[k1]

        # A：発注側に001が無いヘッダは全行 Conflicts
        mask_A_bad = ~fu_hdr.isin(p_has001)

        # B：受注側が一部存在するヘッダでは「受注側に無い行」を Conflicts（受注ゼロ=ルールCは落とさない）
        mask_B_need_order = fu_hdr.isin(o_headers)
        fu_tuples = pd.Series(list(zip(fu_hdr.tolist(), fu_ln.tolist(), strict=False)), index=fu_hdr.index)
        mask_B_missing_in_order = mask_B_need_order & ~fu_tuples.isin(o_keyset)

        mask_immediate_conf = (mask_A_bad | mask_B_missing_in_order).reindex(final_un.index, fill_value=False)

        final_un_conflicts = final_un.loc[mask_immediate_conf].copy()
        final_un_eval = final_un.loc[~mask_immediate_conf].copy()

    if not final_un_conflicts.empty:

        def _reason_for_row(r):
            oid = _nz(r.get(k0))
            ln = _zero_pad3(r.get(k1))
            key_disp = _fmt_key(oid, ln)
            if oid not in p_has001:
                return ('BATCH_STARTS_NOT_001', f'発注側に001行が無いため登録禁止: {key_disp}')
            return ('MISSING_IN_ORDER_CSV', f'受注側に該当行が無いため登録禁止: {key_disp}')

        rs = final_un_conflicts.apply(_reason_for_row, axis=1)
        meta = pd.DataFrame(
            {
                'reason_code': [x[0] for x in rs],
                'reason_detail': [x[1] for x in rs],
                'source': ['p_order_upsert.csv'] * len(final_un_conflicts),
                'timestamp': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * len(final_un_conflicts),
            }
        )
        conflicts.append(pd.concat([meta, final_un_conflicts[p_cols].reset_index(drop=True)], axis=1))

    final_un = final_un_eval

    if not final_un.empty:
        ctx = build_acceptance_context(activeDf, allComp, final_un, keyCols=keyCols, order_csv=order_csv)

        # R1/R2/R4 を評価（理由も付与）
        eval_results = final_un.apply(lambda r: _evaluate_accept_row(r, ctx, keyCols=keyCols), axis=1)
        final_un = final_un.copy()
        # タプル分解
        tmp = pd.DataFrame(eval_results.tolist(), columns=['_accept_pre', '_reason_pre', '_detail_pre'])
        final_un = pd.concat([final_un.reset_index(drop=True), tmp], axis=1)

        # R3: 冪等（DBと“実質同内容”）
        def _r3_idem(row):
            if not row['_accept_pre']:
                return False
            return _is_idempotent_same_as_db(row, ctx, p_cols, keyCols=keyCols)

        final_un['_idem'] = final_un.apply(_r3_idem, axis=1)

        # 最終判定と理由整形
        final_un['_accept'] = final_un['_accept_pre'] & (~final_un['_idem'])

        def _final_reason(r):
            if r['_accept_pre'] and r['_idem']:
                return 'IDEMPOTENT_DUP'
            return r['_reason_pre']

        def _final_detail(r):
            k0, k1 = keyCols
            if r['_accept_pre'] and r['_idem']:
                return f'DBと同内容のためスキップ: {_fmt_key(_nz(r.get(k0)), _zero_pad3(r.get(k1)))}'
            return r['_detail_pre']

        final_un['_reason'] = final_un.apply(_final_reason, axis=1)
        final_un['_detail'] = final_un.apply(_final_detail, axis=1)

        # 分割
        accept_df = final_un.loc[final_un['_accept']].drop(columns=['_accept_pre', '_reason_pre', '_detail_pre', '_idem', '_accept', '_reason', '_detail'], errors='ignore').copy()
        reject_df = final_un.loc[~final_un['_accept']].copy()

        # ACCEPT は既存どおり extra_safe_parts に積む
        if not accept_df.empty:
            for _c in ['受注得意先', '受注納品先', '受注仕入先', '受注商品']:
                if _c not in accept_df.columns:
                    accept_df[_c] = ''
            accept_df['matchMethod'] = 'NO_DB_MATCH_ACCEPTED'
            extra_safe_parts.append(accept_df[p_cols + (['matchMethod'] if 'matchMethod' in accept_df.columns else [])])

        # REJECT は reason_detail を含めて Conflicts へ
        if not reject_df.empty:
            rej = reject_df[p_cols].copy()
            meta = pd.DataFrame(
                {
                    'reason_code': reject_df['_reason'].fillna('NO_DB_MATCH').values,
                    'reason_detail': reject_df['_detail'].fillna('').values,
                    'source': ['p_order_upsert.csv'] * len(reject_df),
                    'timestamp': [datetime.now().strftime('%Y-%m-%d %H:%M:%S')] * len(reject_df),
                }
            )
            conflicts.append(pd.concat([meta, rej.reset_index(drop=True)], axis=1))
    else:
        # final_un が空なら何もしない
        pass

    # --- SAFE 結合 → マスタ結合 → 競合列解消 ---
    safe_parts = [x for x in [m1_safe, m2_safe, db_row_safe, db_prod_safe] if not x.empty]
    if extra_safe_parts:
        safe_parts += [df for df in extra_safe_parts if not df.empty]
    safe_parts = [df.loc[:, ~df.columns.duplicated()].copy() for df in safe_parts]
    mergedActive = pd.concat(safe_parts, ignore_index=True) if safe_parts else pd.DataFrame(columns=p_cols)

    # 安全対策: マスタ結合キー列が無いと KeyError で落ちるので事前保証（空でOK）
    mergedActive = _ensure_cols(mergedActive, ['受注得意先', '受注納品先', '受注仕入先'], fill_value='')

    mergedActive = mergedActive.merge(client_csv, left_on='受注得意先', right_on='得意先', how='left').merge(delivery_csv, left_on='受注納品先', right_on='納品先', how='left').merge(supplier_csv, left_on='受注仕入先', right_on='仕入先', how='left')
    mergedActive.drop(columns=['得意先', '納品先', '仕入先'], errors='ignore', inplace=True)
    mergedActive = resolve_column_conflicts(mergedActive)

    activeDf = dedupeCols(activeDf)
    activeDf = activeDf.loc[:, ~activeDf.columns.duplicated()]
    mergedActive = mergedActive.loc[:, ~mergedActive.columns.duplicated()]

    # CSV に無い手動列を保持
    csvColsUnion = set()
    for src in (order_csv, porder_csv, client_csv, delivery_csv, supplier_csv):
        csvColsUnion |= set(src.columns) - set(keyCols)
    manualCols = [c for c in origCols if c not in keyCols and c not in csvColsUnion and c not in ('mergeFlag', 'matchMethod')]
    manualDf = activeDf.set_index(list(keyCols))[manualCols].copy()

    # --- 再分割 & アーカイブ確定反映 ---
    combined = pd.concat([activeDf, mergedActive, compNew], ignore_index=True).drop_duplicates(subset=list(keyCols), keep='last')

    newComp, newActive = splitCompletedActive(combined)

    # （早期除外）：Completedに存在するキーは Active から落とす（復帰はGAS担当）
    # この時点では finalComp はまだ無いので、現時点の allComp（= compOld + compNew）を使う
    comp_keys_now = _dedupe_key_df(allComp, keyCols)
    before_n2 = len(newActive)
    newActive = _filter_out_keys(newActive, comp_keys_now, keyCols)
    dropped2 = before_n2 - len(newActive)
    if dropped2 > 0:
        print(f'[SAFE] Dropped {dropped2} rows from newActive because keys exist in Completed (snapshot).')

    date_src_B = newComp if not newComp.empty else date_src_A
    archive_sh = ensure_archive_and_maybe_rotate(client, archive_state, rows_to_add=len(newComp), date_src_df=date_src_B)

    compWs = ensure_completed_sheet(
        archive_sh,
        min_rows=max(2, len(newComp) + 5),
        min_cols=max(10, len(dbDf.columns) + 5),
    )

    # ★重要：書き込み直前に再読込（将来GASがCompletedを触っても消しにくくする）
    compOld2 = add_missing_columns(getSheetDf(compWs), list(keyCols))

    # DBスキーマ優先で列順を確定（既存Completed固有列は末尾に保持）
    expCols2 = origCols + [c for c in compOld2.columns if c not in origCols]

    frames = []
    if not compOld2.empty:
        frames.append(compOld2.reindex(columns=expCols2, fill_value=''))

    if not compNew.empty:
        frames.append(compNew.reindex(columns=expCols2, fill_value=''))

    if not newComp.empty:
        frames.append(newComp.reindex(columns=expCols2, fill_value=''))

    if frames:
        finalComp = pd.concat(frames, ignore_index=True).drop_duplicates(subset=list(keyCols), keep='last').sort_values(by=list(keyCols), ascending=True).reset_index(drop=True)
    else:
        finalComp = pd.DataFrame(columns=expCols2).copy()

    for col in finalComp.columns:
        finalComp[col] = finalComp[col].map(cleanValue)

    clearAndWriteDf(compWs, finalComp)

    # --- Active 書き戻し（手動列を復元） ---
    newActive_idx = newActive.set_index(list(keyCols))
    newActive_idx = newActive_idx.drop(columns=manualCols, errors='ignore').join(manualDf, how='left')
    finalActive = newActive_idx.reset_index()

    if 'ステータス' not in finalActive.columns:
        finalActive['ステータス'] = ''
    finalActive['ステータス'] = finalActive['ステータス'].fillna('').astype(str).str.strip()
    finalActive.loc[finalActive['ステータス'] == '', 'ステータス'] = '船積み前'

    for col in finalActive.columns:
        finalActive[col] = finalActive[col].map(cleanValue)

    newCols = [c for c in finalActive.columns if c not in origCols]
    finalCols = origCols + newCols
    finalActive = finalActive.sort_values(by=list(keyCols), ascending=True).reset_index(drop=True)

    # 安全対策（最重要）：Completedに存在するキーはDBへ戻さない（復帰はGAS担当）
    # 書き込み済みの finalComp からキーを取るのが最も確実（compWsを再読込してもOK）
    comp_keys_latest = _dedupe_key_df(finalComp, keyCols)

    # （監査）：Active と Completed のキー重複を検知（※この時点ではまだ除外前なのでWARNのみ）
    a_keys = _df_keys(finalActive, keyCols).drop_duplicates()
    c_keys = _df_keys(finalComp, keyCols).drop_duplicates()
    dup = a_keys.merge(c_keys.assign(_dup=True), on=list(keyCols), how='inner')
    if not dup.empty:
        sample = dup.head(5).apply(lambda r: _fmt_key(r[keyCols[0]], r[keyCols[1]]), axis=1).tolist()
        print(f'[WARN] {len(dup)} duplicate keys between Active and Completed. e.g. {sample}')

    # Completedキーを Active から除外（復活はGAS担当）
    before_n = len(finalActive)
    finalActive = _filter_out_keys(finalActive, comp_keys_latest, keyCols)
    dropped = before_n - len(finalActive)
    if dropped > 0:
        print(f'[SAFE] Dropped {dropped} rows from Active because keys exist in Completed.')

    # -----------------------------
    # キー保存性チェック（FATAL停止）
    # 目的：元DBにあったキーが Active/Completed のどちらにも無い = 消失 を絶対に防ぐ
    # -----------------------------
    db_keys_all = _dedupe_key_df(dbDf, keyCols)  # 元DB（読み込み時点）
    a_keys_all = _dedupe_key_df(finalActive, keyCols)  # 除外後Active
    c_keys_all = _dedupe_key_df(finalComp, keyCols)  # Completed最終

    ac_union = pd.concat([a_keys_all, c_keys_all], ignore_index=True).drop_duplicates()
    missing_from_both = _left_only_keys(db_keys_all, ac_union, keyCols)

    if not missing_from_both.empty:
        sample = missing_from_both.head(20).apply(lambda r: _fmt_key(r[keyCols[0]], r[keyCols[1]]), axis=1).tolist()
        raise RuntimeError(f'[FATAL] 元DBに存在したキーが Active/Completed のどちらにも存在しません（消失検知）: {len(missing_from_both)} keys. sample={sample}')

    # 追加の方針チェック（推奨）：元DBで「完了」だったキーは Completed に必ず存在
    if 'ステータス' in dbDf.columns:
        db_done = dbDf.copy()
        st = db_done['ステータス'].fillna('').astype(str).str.strip()
        db_done_keys = _dedupe_key_df(db_done.loc[st.eq('完了')], keyCols)
        missing_done_in_comp = _left_only_keys(db_done_keys, c_keys_all, keyCols)

        if not missing_done_in_comp.empty:
            sample = missing_done_in_comp.head(20).apply(lambda r: _fmt_key(r[keyCols[0]], r[keyCols[1]]), axis=1).tolist()
            raise RuntimeError(f'[FATAL] 元DBで完了だったキーが Completed に存在しません（完了退避失敗）: {len(missing_done_in_comp)} keys. sample={sample}')

    # 追加の方針チェック（推奨）：Completed のキーが Active に残っていないこと（復活禁止）
    dup_after = a_keys_all.merge(c_keys_all.assign(_dup=True), on=list(keyCols), how='inner')
    if not dup_after.empty:
        sample = dup_after.head(20).apply(lambda r: _fmt_key(r[keyCols[0]], r[keyCols[1]]), axis=1).tolist()
        raise RuntimeError(f'[FATAL] Completed のキーが Active に残っています（復活禁止違反）: {len(dup_after)} keys. sample={sample}')

    # --- ここまでOKならDBへ書く ---
    clearAndWriteDf(dbWs, finalActive[finalCols].copy())
    print('DBシート（Active）を SAFE のみで更新しました。')

    # --- Conflicts を書き出し（追記モード：メタ＋DB列（可変）＋CSV-only） ---
    confWs = ensure_conflicts_sheet(sh, min_rows=2, min_cols=10)

    # メタ列（先頭固定）
    meta_cols = ['reason_code', 'reason_detail', 'source', 'timestamp']

    # DBヘッダ（可変）：origCols をベースに（メタ列は除外）
    base_cols = [c for c in list(origCols) if c not in meta_cols]

    # CSV 由来の列（order/p_order のユニオン）から、DB列とメタを除外
    csv_cols_union = sorted(set(p_cols) | set(o_cols))
    csv_only_cols = [c for c in csv_cols_union if c not in base_cols and c not in meta_cols]

    # 今回の標準列順（目安）
    standard_cols = meta_cols + base_cols + csv_only_cols

    # 既存 Conflicts を読み込み（空なら空DF）
    _vals = confWs.get_all_values()
    if _vals:
        existing_header = _vals[0]
        existing_body = _vals[1:] if len(_vals) > 1 else []
        existing_df = pd.DataFrame(existing_body, columns=existing_header) if existing_body else pd.DataFrame(columns=existing_header)
    else:
        existing_df = pd.DataFrame(columns=standard_cols)

    if conflicts:
        # 今回検出分を結合
        conf_df = pd.concat(conflicts, ignore_index=True, sort=False)

        # 列ユニオン（既存列を優先 → 標準列 → 今回新規列）
        union_cols = list(dict.fromkeys(list(existing_df.columns) + standard_cols + list(conf_df.columns)))

        # 不足列は空で補完し、列順をそろえる
        existing_df = existing_df.reindex(columns=union_cols, fill_value='')
        conf_df = conf_df.reindex(columns=union_cols, fill_value='')

        # 値整形
        for col in conf_df.columns:
            conf_df[col] = conf_df[col].map(cleanValue)

        # 追記（下に積む）
        appended = pd.concat([existing_df, conf_df], ignore_index=True)

        # 書き戻し（ヘッダ＋全件）
        clearAndWriteDf(confWs, appended)
        print('Conflicts シートを追記更新しました（既存＋今回検出）')
    else:
        # 追加が無い場合は何もしない（既存を保持）
        print('Conflicts 追加なし（既存は保持）')


def collapse_csv_suffix(df: pd.DataFrame, prefer: str = 'csv') -> pd.DataFrame:
    """
    *_csv と素の列を併合（coalesce）し、*_csv を削除する。
    prefer = "csv" | "base" で、欠損時の補完優先を選べる。
        - "csv": baseが欠損のとき*_csv値で埋める（CSV優先）
        - "base": *_csvが欠損のときbase値で埋める（DB優先）
    """
    out = df.copy()
    csv_cols = [c for c in out.columns if c.endswith('_csv')]
    for c in csv_cols:
        base = c[:-4]
        if base in out.columns:
            if prefer == 'csv':
                out[base] = out[base].combine_first(out[c])
            else:
                out[base] = out[c].combine_first(out[base])
            out.drop(columns=[c], inplace=True)
        else:
            # 衝突相手が無い場合は列名を素に戻して未知列化を防止
            out.rename(columns={c: base}, inplace=True)
    return out


# --------------------------------------------------------------------------------------
# DataFrame/列ユーティリティ
# --------------------------------------------------------------------------------------


def clearAndWriteDf(sheet, df, chunkCellLimit=10000):
    """シートをクリアして DF をヘッダ付きで書き込み（大きい場合は分割）。"""
    from gspread.utils import rowcol_to_a1

    for col in df.select_dtypes(include=['category']).columns:
        df[col] = df[col].astype(object)
    df = df.where(pd.notnull(df), '')

    rows = df.shape[0] + 1
    cols = df.shape[1]

    try:
        sheet.resize(rows=max(rows, sheet.row_count), cols=max(cols, sheet.col_count))
    except APIError as e:
        print(f'[Warning] Sheet resize failed: {e}')

    if cols == 0 or rows <= 1:
        sheet.clear()
        return

    sheet.clear()

    sheet.format('A:ZZZ', {'numberFormat': {'type': 'TEXT'}})

    data = [list(df.columns)] + df.values.tolist()
    total_cells = len(data) * cols

    if total_cells <= chunkCellLimit:
        end_cell = rowcol_to_a1(len(data), cols)
        safe_update(sheet, range_name=f'A1:{end_cell}', values=data)
    else:
        safe_bulk_update(sheet, start_cell='A1', values=data)


def splitCompletedActive(df: pd.DataFrame):
    """
    ステータスで Completed/Active に分割する（完了区分は見ない）。
    - ステータス == '完了' の行のみ Completed
    - それ以外は Active
    """
    if 'ステータス' not in df.columns:
        return pd.DataFrame(columns=df.columns), df.copy()

    tmp = df.copy()
    status = tmp['ステータス'].fillna('').astype(str).str.strip()
    done_mask = status == '完了'

    done = tmp[done_mask].copy()
    active = tmp[~done_mask].copy()

    # Completed に渡す行はステータス=完了を保証
    done['ステータス'] = '完了'
    return done, active


# def splitCompletedActive(
#     df: pd.DataFrame,
#     *,
#     active_exception_supplier_keywords: tuple[str, ...] = (
#         'before ship',
#         '7days',
#     ),
# ):
#     """
#     ステータス/完了区分を見て Completed/Active に分割する。
#     追加仕様:
#     - 発注仕入先略称 に active_exception_supplier_keywords の
#         いずれかが【部分一致】する行は、
#         完了区分が完了でも Active に残す。
#     - 大文字小文字は区別しない。
#     """
#     flags = ['ステータス', '受注売上完了区分', '発注仕入完了区分']
#     if not any(f in df.columns for f in flags):
#         return pd.DataFrame(columns=df.columns), df.copy()

#     tmp = df.copy()

#     # --- 完了判定（値は書き換えずマスクのみ）---
#     status = tmp['ステータス'].fillna('').astype(str).str.strip() if 'ステータス' in tmp.columns else pd.Series('', index=tmp.index)

#     done_mask = status == '完了'

#     for flag in ['受注売上完了区分', '発注仕入完了区分']:
#         if flag in tmp.columns:
#             done_mask = done_mask | (tmp[flag].fillna('').astype(str).str.strip() == '1:完了')

#     # --- 例外：発注仕入先略称の部分一致 ---
#     if '発注仕入先略称' in tmp.columns and active_exception_supplier_keywords:
#         abbr = tmp['発注仕入先略称'].fillna('').astype(str)

#         exception_mask = pd.Series(False, index=tmp.index)
#         for kw in active_exception_supplier_keywords:
#             if kw:
#                 exception_mask = exception_mask | abbr.str.contains(kw, case=False, regex=False)

#         # 例外は Completed から除外 → Active に残す
#         done_mask = done_mask & ~exception_mask

#     done = tmp[done_mask].copy()
#     active = tmp[~done_mask].copy()

#     # --- 追加仕様：Completed に渡す行はステータス=完了を保証 ---
#     if 'ステータス' not in done.columns:
#         done['ステータス'] = '完了'
#     else:
#         done['ステータス'] = '完了'

#     return done, active


# def splitCompletedActive(df: pd.DataFrame):
#     """ステータス/完了区分を見て Completed/Active に分割。"""
#     flags = ['ステータス', '受注売上完了区分', '発注仕入完了区分']
#     if not any(f in df.columns for f in flags):
#         return pd.DataFrame(columns=df.columns), df.copy()
#     tmp = df.copy()
#     if 'ステータス' in tmp.columns:
#         tmp['ステータス'] = tmp['ステータス'].fillna('').str.strip()
#     for flag in ['受注売上完了区分', '発注仕入完了区分']:
#         if flag in tmp.columns:
#             tmp.loc[tmp[flag] == '1:完了', 'ステータス'] = '完了'
#     done = tmp[tmp['ステータス'] == '完了']
#     active = tmp[tmp['ステータス'] != '完了']
#     return done, active


def ensure_columns(df: pd.DataFrame, cols: list[str], fill_value='') -> pd.DataFrame:
    """指定列が無ければ追加して埋める。"""
    for c in cols:
        if c not in df.columns:
            df[c] = fill_value
    return df


def drop_if_exists(df: pd.DataFrame, cols: list[str]) -> pd.DataFrame:
    """存在する列だけ削除。"""
    to_drop = [c for c in cols if c in df.columns]
    return df.drop(columns=to_drop) if to_drop else df


def rename_if_exists(df: pd.DataFrame, rename_map: dict[str, str]) -> pd.DataFrame:
    """存在する列だけリネーム。"""
    valid_map = {old: new for old, new in rename_map.items() if old in df.columns}
    return df.rename(columns=valid_map) if valid_map else df


def to_python_scalar(v):
    """cleanValue の別名（互換用）。"""
    return cleanValue(v)


def cleanValue(val):
    """数値/小数/ゼロ埋めの文字列などを安全に整形。"""
    if pd.isna(val):
        return ''
    if isinstance(val, (np.integer, int)):
        return int(val)
    if isinstance(val, (np.floating, float)):
        v = float(val)
        return int(v) if v.is_integer() else v
    if isinstance(val, str):
        s = val.replace(',', '')
        if re.fullmatch(r'0\d+', s):
            return s
        if re.fullmatch(r'-?[1-9]\d*|0', s):
            return int(s)
        m = re.fullmatch(r'(-?\d+)\.(\d+)', s)
        if m:
            i, f = m.groups()
            return int(i) if set(f) == {'0'} else float(s)
    return val


def loadCsvWithPrefix(path, prefix, excludeCols=None):
    """CSV を読み、指定 prefix を列名へ付与（除外列はそのまま）。"""
    df = pd.read_csv(path, dtype=str, low_memory=False)
    newCols = []
    for c in df.columns:
        if excludeCols and c in excludeCols:
            newCols.append(c)
        elif c.startswith(prefix):
            newCols.append(c)
        else:
            newCols.append(f'{prefix}{c}')
    df.columns = newCols
    return df


def mark_exists(df, keys_df, key_cols):
    """DBに存在するかをベクトルで判定（mergeベース）。"""
    exists = df[list(key_cols)].merge(keys_df[list(key_cols)].drop_duplicates().assign(_exists=True), on=list(key_cols), how='left')['_exists'].fillna(False).astype(bool)
    out = df.copy()
    out['exists_in_db'] = exists
    return out


def dedupeCols(df):
    """同名列がある場合、_x/_y 起源の重複に連番を付けて衝突回避。"""
    seen = {}
    cols = []
    for c in df.columns:
        if c not in seen:
            seen[c] = 0
            cols.append(c)
        else:
            seen[c] += 1
            cols.append(f'{c}{seen[c]}')
    df.columns = cols
    return df


def add_missing_columns(df: pd.DataFrame, cols: list[str], default='') -> pd.DataFrame:
    """既存順を保ちつつ欠け列を後ろに追加。"""
    all_cols = list(df.columns) + [c for c in cols if c not in df.columns]
    return df.reindex(columns=all_cols, fill_value=default)


def resolve_column_conflicts(df, *, csv_prefer: str = 'csv'):
    """
    1) *_x / *_y を base 列へ集約（combine_first）
    2) *_csv を base 列へ畳み込み（なければリネーム）
        csv_prefer: "csv"（CSV優先） | "base"（DB優先）
    戻り値は新DF（元dfは不変）
    """
    from collections import defaultdict

    out = df.copy()

    # 1) *_x / *_y
    conflict_map = defaultdict(list)
    for col in out.columns:
        if col.endswith('_x') or col.endswith('_y'):
            conflict_map[col[:-2]].append(col)

    for base, _ in conflict_map.items():  # B007回避
        col_x, col_y = f'{base}_x', f'{base}_y'
        has_x, has_y = col_x in out.columns, col_y in out.columns
        if has_x and has_y:
            out[base] = out[col_x].combine_first(out[col_y])
        elif has_x:
            out[base] = out[col_x]
        elif has_y:
            out[base] = out[col_y]
    if conflict_map:
        drop_cols = [c for v in conflict_map.values() for c in v]
        out.drop(columns=drop_cols, inplace=True, errors='ignore')

    # 2) *_csv
    csv_cols = [c for c in out.columns if c.endswith('_csv')]
    for c in csv_cols:
        base = c[:-4]
        if base in out.columns:
            if csv_prefer == 'csv':
                out[base] = out[base].combine_first(out[c])
            else:
                out[base] = out[c].combine_first(out[base])
            out.drop(columns=[c], inplace=True, errors='ignore')
        else:
            out.rename(columns={c: base}, inplace=True)

    return out.loc[:, ~out.columns.duplicated()].copy()


def safe_update(sheet, range_name, values, max_retries=6, delay_seconds=5):
    """Worksheet.update をリトライ付きで安全に実行（429は長めに待つ）。"""
    backoff_429 = [30, 60, 120, 180, 240, 300]  # 秒（最大でも5分程度）
    for attempt in range(1, max_retries + 1):
        try:
            sheet.update(range_name=range_name, values=values)
            return
        except APIError as e:
            msg = str(e)
            is_429 = ('[429]' in msg) or ('Quota exceeded' in msg)

            if attempt == max_retries:
                print(f'[safe_update] retry exhausted. last error: {e}')
                raise

            if is_429:
                wait = backoff_429[min(attempt - 1, len(backoff_429) - 1)]
                print(f'[safe_update] 429 quota exceeded. sleep {wait}s (attempt {attempt}/{max_retries})')
                time.sleep(wait)
            else:
                print(f'[safe_update] APIError: {e} (attempt {attempt}/{max_retries}) sleep {delay_seconds}s')
                time.sleep(delay_seconds)


def safe_bulk_update(sheet, start_cell, values, max_retries=3, delay_seconds=5):
    """範囲更新が失敗した場合、行ごと更新にフォールバック。"""
    import re as _re

    from gspread.utils import rowcol_to_a1

    m = _re.match(r'([A-Z]+)(\d+)', start_cell)
    if not m:
        raise ValueError("start_cell must be like 'A1', 'B2', etc.")
    col_letter, start_row = m.groups()
    start_row = int(start_row)
    start_col_idx = ord(col_letter.upper()) - ord('A') + 1
    end_row = start_row + len(values) - 1
    end_col = len(values[0])
    end_cell = rowcol_to_a1(end_row, start_col_idx + end_col - 1)
    full_range = f'{start_cell}:{end_cell}'

    try:
        safe_update(sheet, full_range, values, max_retries=max_retries, delay_seconds=delay_seconds)
    except Exception:
        print('[Fallback] Bulk update failed. Switching to row-wise updates.')
        for i, row in enumerate(values):
            cell = rowcol_to_a1(start_row + i, start_col_idx)
            try:
                safe_update(sheet, cell, [row], max_retries=max_retries, delay_seconds=delay_seconds)
            except Exception as e2:
                print(f'[Error] Failed to write row {i + 1} to {cell}: {e2}')


def build_fact_pk(df: pd.DataFrame) -> pd.Series:
    """
    FACT_売上の主キー __pk を生成: 売上NO|行NO
    - 両列とも文字列固定（ゼロ詰めはCSV側前提、念のためstripは実施）
    """
    if '売上NO' not in df.columns or '行NO' not in df.columns:
        raise MyError("FACT_売上CSVに '売上NO' と '行NO' が必要です。")

    a = df['売上NO'].astype(str).str.strip()
    b = df['行NO'].astype(str).str.strip()
    return a + '|' + b


def compute_row_hash(df: pd.DataFrame, *, exclude_cols: list[str] | None = None) -> pd.Series:
    """
    行単位の安定hash（md5）。
    - 文字列固定で比較できるよう、NaNは空文字、区切り文字で連結してhash化
    """
    exclude = set(exclude_cols or [])
    cols = [c for c in df.columns if c not in exclude]

    # 連結用の区切り（通常データに混入しにくい）
    sep = '\u241f'  # Unit Separator 記号

    def _hash_row(row_vals) -> str:
        s = sep.join('' if pd.isna(v) else str(v) for v in row_vals)
        return hashlib.md5(s.encode('utf-8')).hexdigest()

    # values: ndarrayで高速化
    vals = df[cols].values
    return pd.Series((_hash_row(r) for r in vals), index=df.index, dtype=str)


def ensure_fact_internal_cols(df: pd.DataFrame) -> pd.DataFrame:
    """
    dfに __pk と __hash を付与して返す（末尾列）。
    """
    out = df.copy()
    out['__pk'] = build_fact_pk(out)
    out['__hash'] = compute_row_hash(out, exclude_cols=['__pk', '__hash'])
    return out


def colnum_to_letter(n: int) -> str:
    """1->A, 27->AA"""
    s = ''
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def read_pk_hash_index(ws) -> dict[str, tuple[int, str]]:
    """
    FACT_売上シートから __pk / __hash のみ読み、 {pk: (rownum, hash)} を返す。
    rownum はシート上の実行行番号（2行目=データ先頭）。
    """
    header = ws.row_values(1)
    if not header:
        return {}

    try:
        pk_idx = header.index('__pk') + 1  # 1-based
        hs_idx = header.index('__hash') + 1
    except ValueError:
        # まだ列が無い場合は空
        return {}

    c1 = colnum_to_letter(min(pk_idx, hs_idx))
    c2 = colnum_to_letter(max(pk_idx, hs_idx))
    rng = f'{c1}2:{c2}'  # 下端省略→filledまで返る

    rows = ws.get(rng)  # 1リクエスト
    out: dict[str, tuple[int, str]] = {}

    # 取得レンジが pk/hash の順と一致するとは限らないので並びを補正
    pk_pos = 0 if pk_idx <= hs_idx else 1
    hs_pos = 1 - pk_pos

    for i, r in enumerate(rows, start=2):
        pk = (r[pk_pos] if len(r) > pk_pos else '').strip()
        hs = (r[hs_pos] if len(r) > hs_pos else '').strip()
        if pk:
            out[pk] = (i, hs)
    return out


def group_consecutive_rows(rows: list[int]) -> list[tuple[int, int]]:
    """
    [2,3,4,10,11] -> [(2,4),(10,11)]
    """
    if not rows:
        return []
    rows = sorted(set(rows))
    out = []
    start = prev = rows[0]
    for r in rows[1:]:
        if r == prev + 1:
            prev = r
            continue
        out.append((start, prev))
        start = prev = r
    out.append((start, prev))
    return out


def apply_fact_diff(ws, df_new: pd.DataFrame, *, cells_per_batch: int = 200000):
    """
    FACT_売上へ差分反映。
    - df_new は __pk/__hash を含む想定（ensure_fact_internal_cols済み）
    """
    # 1) シートヘッダを確定（無ければ新規ヘッダで初期化）
    header = ws.row_values(1)
    if not header:
        header = list(df_new.columns)
        end_col = colnum_to_letter(len(header))
        safe_update(ws, f'A1:{end_col}1', [header])
    else:
        # df側に列追加があれば末尾追加（既存順は維持）
        missing = [c for c in df_new.columns if c not in header]
        if missing:
            header = header + missing
            end_col = colnum_to_letter(len(header))
            safe_update(ws, f'A1:{end_col}1', [header])

    # 2) df列をシート列順にそろえる
    df_aligned = df_new.reindex(columns=header, fill_value='')

    # 3) 既存索引（pk->(row,hash)）
    idx = read_pk_hash_index(ws)

    # 4) 判定（update / append / skip）
    updates: dict[int, list[str]] = {}  # rownum -> row_values
    appends: list[list[str]] = []

    for _, row in df_aligned.iterrows():
        pk = str(row.get('__pk', '')).strip()
        hs = str(row.get('__hash', '')).strip()
        if not pk:
            continue

        if pk not in idx:
            appends.append([cleanValue(v) for v in row.tolist()])
            continue

        rownum, old_hash = idx[pk]
        if hs and old_hash == hs:
            continue  # 変更なし

        updates[rownum] = [cleanValue(v) for v in row.tolist()]

    # 5) 更新（連続行をまとめてレンジ更新）
    if updates:
        rows_sorted = sorted(updates.keys())
        ranges = group_consecutive_rows(rows_sorted)

        ncols = len(header)
        col_end = colnum_to_letter(ncols)

        for r1, r2 in ranges:
            block = [updates[r] for r in range(r1, r2 + 1)]
            rng = f'A{r1}:{col_end}{r2}'
            safe_update(ws, rng, block)

    # 6) 追加（末尾にまとめて追記）
    if appends:
        # 末尾行（__pk列が埋まっている最終行＋1を推定）
        # idxが空でないなら最大row+1、空なら2行目
        start_row = max((r for r, _ in idx.values()), default=1) + 1

        ncols = len(header)
        col_end = colnum_to_letter(ncols)

        # cells_per_batchでチャンク分割（update回数を抑えるため大きめ推奨）
        max_rows_per_chunk = max(1, cells_per_batch // max(1, ncols))

        cursor = 0
        while cursor < len(appends):
            chunk = appends[cursor : cursor + max_rows_per_chunk]
            r1 = start_row + cursor
            r2 = r1 + len(chunk) - 1
            rng = f'A{r1}:{col_end}{r2}'
            safe_update(ws, rng, chunk)
            cursor += len(chunk)


def sync_fact_sales_diff_from_csv(client: gspread.Client, spreadsheet_id: str, files_dir: str):
    """
    sales_upsert.csv -> FACT_売上 を差分更新（__pk/__hash 追加前提）。
    """
    ss = client.open_by_key(spreadsheet_id)
    fact_path = os.path.join(files_dir, 'sales_upsert.csv')

    df_fact = load_csv_as_text(fact_path)
    df_fact = normalize_dates_textual(df_fact)
    df_fact = ensure_fact_internal_cols(df_fact)

    ws_fact = ensure_named_sheet(
        ss,
        'FACT_売上',
        min_rows=max(2, len(df_fact) + 10),
        min_cols=max(10, df_fact.shape[1] + 2),
    )

    apply_fact_diff(ws_fact, df_fact, cells_per_batch=200000)
    print(f'[sync_fact_sales_diff_from_csv] FACT_売上: input={len(df_fact)} rows, updated/append applied')


# --------------------------------------------------------------------------------------
# 売上・得意先
# --------------------------------------------------------------------------------------
def _type_absolute_date(ymd: str):
    """'yyyy/MM/dd' または 'yyyy-MM-dd' を '年→Enter→月→Enter→日' でタイプする。"""
    ymd = ymd.strip().replace('-', '/')
    parts = ymd.split('/')
    if len(parts) != 3:
        raise MyError(f'ERROR: 不正な日付文字列です: {ymd}')
    y, m, d = parts
    ag.typewrite(y)
    ag.press('enter')
    ag.typewrite(m.zfill(2))
    ag.press('enter')
    ag.typewrite(d.zfill(2))


def export_timelimit_absolute(window_title: str, start_ymd: str, end_ymd: str, retries: int = 5, delay_sec: float = 5.0):
    """
    明細出力（絶対期間）。保存完了（CSV存在）まで最大 retries 回リトライ。
    チェック場所: get_base_path() 直下。存在のみ判定。失敗時はログしてスキップ。
    """

    def _do_once():
        copy_paste(window_title)
        ag.press('enter', presses=2)

        app_check(window_title, 1)
        ag.press('enter')
        _type_absolute_date(start_ymd)  # 開始日
        ag.press('enter')
        _type_absolute_date(end_ymd)  # 終了日
        ag.press('f1')

        app_check(window_title + '問合せ', 10)  # 長期間検索の想定
        ag.press('f1')

        # 保存先フォルダ選択
        app_check('出力', 1)
        ag.hotkey('ctrlleft', 'f')
        app_check('名前を付けて保存', 1)

        base_path_env = get_base_path_for_env()
        ag.press('left')
        copy_paste(base_path_env + '\\')
        ag.press('enter')
        app_wait('名前を付けて保存')

        app_check('出力', 1)
        ag.press('1')
        ag.press('enter')
        ag.press('2')
        ag.press('enter')
        ag.press('1')
        ag.press('enter')
        ag.press('0')
        ag.press('enter')
        ag.press('1')
        ag.press('enter')
        ag.press('1')
        ag.press('f1')

        app_check('情報', 10)  # 長期間検索の想定
        ag.press('enter')

        app_check(window_title + '問合せ', 1)
        ag.press('f12')
        app_check(window_title, 1)
        ag.press('f12')

        img_check(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'image', '001_削除.png'), 0, 0, 1)
        ag.press('tab', presses=2)

    keyword = f'{window_title}問合せ'
    base_dir_chk = get_base_path()

    for attempt in range(1, retries + 1):
        _do_once()
        found = _find_csv_by_keyword(base_dir_chk, keyword)
        if found:
            print(f'[export_timelimit_absolute] 保存確認OK: {found}')
            return True
        else:
            if attempt < retries:
                print(f'[export_timelimit_absolute] CSV未検出（{keyword}）。{delay_sec:.0f}秒後にリトライ {attempt}/{retries}')
                time.sleep(delay_sec)
            else:
                print(f'[export_timelimit_absolute] CSV未検出（{keyword}）。最大リトライ到達のためスキップ')
                return False


def load_csv_as_text(path: str) -> pd.DataFrame:
    """CSVを dtype=str で読み込み、NaN→''、すべて文字列化（ゼロ詰め保持）。"""
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    df = pd.read_csv(path, dtype=str, low_memory=False)
    df = df.fillna('')
    # すべて文字列へ（既にstrだが、念のため統一）
    for c in df.columns:
        df[c] = df[c].astype(str)
    return df


def normalize_dates_textual(df: pd.DataFrame) -> pd.DataFrame:
    """
    列名に「日」を含む列のみ、'yyyy/MM/dd' の文字列へ正規化（パース不可は原文のまま）。
    pandasの自動推測は使わず、警告を出さない実装。
    """
    if df.empty:
        return df

    from datetime import datetime

    def _parse_to_ymd(s: str) -> str:
        if not s:
            return s
        t = str(s).strip()
        if not t:
            return t

        # 統一のために区切りを '/' に寄せる
        t_std = t.replace('-', '/').replace('.', '/').replace('年', '/').replace('月', '/').replace('日', '')

        # 候補フォーマット（増やす場合はここに追加）
        candidates = [
            '%Y/%m/%d',
            '%Y/%m/%d',  # 二重だが順序明示のため残す
        ]

        # 8桁の連番（yyyyMMdd）にも対応
        digits = ''.join(ch for ch in t if ch.isdigit())
        if len(digits) == 8 and digits.isdigit():
            try:
                dt = datetime.strptime(digits, '%Y%m%d')
                return dt.strftime('%Y/%m/%d')
            except Exception:
                pass

        for fmt in candidates:
            try:
                dt = datetime.strptime(t_std, fmt)
                return dt.strftime('%Y/%m/%d')
            except Exception:
                continue

        # "yyyy/m/d" のように0詰めの無い形式にも対応（都度分解）
        if '/' in t_std:
            parts = t_std.split('/')
            if len(parts) == 3 and all(parts[i].strip().isdigit() for i in range(3)):
                try:
                    y = int(parts[0].strip())
                    m = int(parts[1].strip())
                    d = int(parts[2].strip())
                    dt = datetime(year=y, month=m, day=d)
                    return dt.strftime('%Y/%m/%d')
                except Exception:
                    pass

        # どれにも当たらなければそのまま返す
        return t

    out = df.copy()
    date_like_cols = [c for c in out.columns if '日' in str(c)]
    for col in date_like_cols:
        out[col] = out[col].astype(str).map(_parse_to_ymd)
    return out


def write_df_as_text(sheet, df: pd.DataFrame, cells_per_batch: int = 20000):
    """
    シートをクリアして DF（全て文字列）をヘッダ付きで書き込み。
    行バッチに分割し、1,000行超も確実に全件上書きする。
    """
    from gspread.utils import rowcol_to_a1

    if df is None:
        df = pd.DataFrame()

    # すべて文字列化＆空埋め
    df = df.fillna('')
    for c in df.columns:
        df[c] = df[c].astype(str)

    # シート初期化
    total_rows = df.shape[0]
    total_cols = max(1, df.shape[1])
    try:
        sheet.resize(rows=max(total_rows + 1, sheet.row_count), cols=max(total_cols, sheet.col_count))
    except APIError as e:
        print(f'[Warning] Sheet resize failed: {e}')

    sheet.clear()
    if total_cols == 0:
        return  # ヘッダ無しなら空クリアのみ

    # バッチサイズ計算（セル数基準）
    # 1行あたり total_cols セルなので、cells_per_batch / total_cols 行を一括で投下
    rows_per_batch = max(1, cells_per_batch // total_cols)

    # 1) ヘッダ
    header = [list(df.columns)]
    end_cell = rowcol_to_a1(1, total_cols)
    safe_update(sheet, range_name=f'A1:{end_cell}', values=header)

    # 2) 本体を行バッチで
    start_row = 2
    n = total_rows
    i = 0
    while i < n:
        j = min(i + rows_per_batch, n)
        batch_values = df.iloc[i:j, :].values.tolist()
        end_cell = rowcol_to_a1(start_row + (j - i) - 1, total_cols)
        safe_update(sheet, range_name=f'A{start_row}:{end_cell}', values=batch_values)
        start_row += j - i
        i = j

    # 最終的にサイズをピタッと合わせる（余剰行/列があれば縮める）
    try:
        sheet.resize(rows=total_rows + 1, cols=total_cols)
    except APIError as e:
        print(f'[Warning] Sheet final resize failed: {e}')


def sync_fact_and_dim_from_csvs(client: gspread.Client, spreadsheet_id: str, files_dir: str):
    """
    CSV -> FACT/DIM 同期
    - FACT_売上: 差分（PK+hash）
    - DIM_得意先: 当面フル（必要なら後で差分化）
    """
    ss = client.open_by_key(spreadsheet_id)

    # FACT_売上（差分）
    try:
        sync_fact_sales_diff_from_csv(client, spreadsheet_id, files_dir)
    except FileNotFoundError:
        print('[sync_fact_and_dim_from_csvs] sales_upsert.csv が見つからないため FACT_売上 をスキップ')

    # DIM_得意先（フル）
    dim_path = os.path.join(files_dir, 'client_upsert.csv')
    try:
        df_dim = load_csv_as_text(dim_path)
        df_dim = normalize_dates_textual(df_dim)

        ws_dim = ensure_named_sheet(
            ss,
            'DIM_得意先',
            min_rows=max(2, len(df_dim) + 1),
            min_cols=max(10, df_dim.shape[1] + 2),
        )
        # フル上書きでも update 回数を減らす
        write_df_as_text(ws_dim, df_dim, cells_per_batch=200000)
        print(f'[sync_fact_and_dim_from_csvs] DIM_得意先: {len(df_dim)} 行を書き込みました')
    except FileNotFoundError:
        print('[sync_fact_and_dim_from_csvs] client_upsert.csv が見つからないため DIM_得意先 をスキップ')


def main():
    start = time.time()
    resume_point = 0
    while time.time() - start <= 60 * 3:  # タイムアウト
        try:
            ao_login()
            break
        # except Exception as e:
        except MyError:
            process_close('mstsc.exe')
            traceback.print_exc()
            time.sleep(1)
    if time.time() - start >= 60 * 3:  # タイムアウトした場合
        process_close('mstsc.exe')
        sys.exit(1)

    while time.time() - start <= 60 * 5:  # タイムアウト
        try:
            if resume_point <= 1:
                resume_point = 1
                ao_action()

            if resume_point <= 2:
                resume_point = 2
                move_csv_files_from_base(['受注明細問合せ', '発注明細問合せ', '売上明細問合せ', '得意先マスタ一覧表問合せ', '納品先マスタ一覧表問合せ', '仕入先マスタ一覧表問合せ'])
                # バックアップフォルダ内の30日以上経過したファイルを削除
                delete_old_files(os.path.join(os.path.dirname(sys.argv[0]), 'files', 'csv'))

            # if resume_point <= 3:
            #     resume_point = 3
            #     # "受注明細問合せ"のCSVデータの処理とバックアップ
            #     data1_date_columns = ["受注日", "指定納期", "製造日", "出荷予定日", "納期回答日"]
            #     data1_amount_columns = ["入数", "ｹｰｽ", "受注数", "売上数", "受注残数量", "受注単価", "受注金額", "受注残金額", "原価単価", "原価金額", "粗利金額", "単重", "換算重量", "残換算重量", "手配入数", "手配ｹｰｽ", "手配数", "上代単価"]
            #     data1_decimal_point_columns = ["原価単価", "原価金額", "粗利金額", "単重", "換算重量", "残換算重量", "手配入数", "手配ｹｰｽ", "上代単価"]
            #     data1_number_columns = ["受注管理NO"]
            #     process_data("受注管理NO含む問合せ", data1_date_columns, data1_amount_columns, data1_decimal_point_columns, data1_number_columns)

            #     # "売上明細問合せ"のCSVデータの処理とバックアップ
            #     data2_date_columns = ["売上日", "出荷日", "製造日", "入力回収予定日", "分割回収予定日1", "分割回収予定日2", "分割回収予定日3", "ETD納期"]
            #     data2_amount_columns = ["入数", "ｹｰｽ", "売上数", "売上単価", "売上金額", "原価単価", "原価金額", "粗利金額", "単重", "換算重量", "上代単価"]
            #     data2_decimal_point_columns = ["入数", "ｹｰｽ", "売上単価", "売上金額", "原価単価", "原価金額", "粗利金額", "単重", "換算重量", "上代単価"]
            #     data2_number_columns = ["売上管理NO"]
            #     process_data("売上管理NO含む問合せ", data2_date_columns, data2_amount_columns, data2_decimal_point_columns, data2_number_columns)

            #     # "得意先マスタ一覧表問合せ"のCSVデータの処理とバックアップ
            #     process_data_client("得意先マスタ一覧表問合せ")

            if resume_point <= 4:
                resume_point = 4

                client = authorizeGspread(credentialsFile)
                sh = client.open_by_key(spreadsheetId)
                dbWs = sh.worksheet(dbSheetName)
                run_import_without_log(client, sh, dbWs, filesDir, keyCols=keyCols)
                sync_fact_and_dim_from_csvs(client, spreadsheetId, filesDir)

            break
        # except Exception as e:xxx
        except MyError:
            ao_login()
            traceback.print_exc()
            time.sleep(1)
    if time.time() - start >= 60 * 5:  # タイムアウトした場合
        process_close('mstsc.exe')


if __name__ == '__main__':
    main()
