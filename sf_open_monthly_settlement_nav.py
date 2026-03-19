# -*- coding: utf-8 -*-
"""
Windows desktop automation for 顺发打单.

This version uses simulated mouse clicks instead of Windows control lookup.

Usage:
    1. Install dependencies:
       pip install pywinauto pywin32 pyautogui
    2. First run on Windows to record click points:
       python sf_open_monthly_settlement_nav.py --record
    3. Later runs can execute directly:
       python sf_open_monthly_settlement_nav.py
"""

import argparse
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

import pyautogui
from pywinauto import Desktop
from pywinauto.keyboard import send_keys


APP_SHORTCUT = r"C:\Users\86138\Desktop\顺发打单.lnk"
WINDOW_TITLE_HINT = "顺发"
POINTS_FILE = Path(__file__).with_name("sf_open_monthly_settlement_nav_points.json")
ACTION_STEPS = ["全部订单", "导出", "提交", "我知道了", "通知", "任务中心", "下载文件"]
CLICK_INTERVAL_SECONDS = 1.2

# 提交后等待时间（秒）
SUBMIT_WAIT_SECONDS = 5

# 文件保存配置
SAVE_CONFIG = {
    "directory": r"C:\Orders\Export",           # 保存目录
    "filename_template": "顺发订单_{date}.xlsx",  # 文件名模板，{date} 会被替换为 YYYYMMDD
    "dialog_timeout": 10,                        # 等待另存为对话框超时(秒)
    "save_button_timeout": 5,                    # 等待保存完成超时(秒)
}


def iter_visible_windows():
    desktop = Desktop(backend="uia")
    for win in desktop.windows():
        try:
            if not win.is_visible():
                continue
            yield win
        except Exception:
            continue


def iter_all_windows():
    """枚举所有窗口（包括不可见的），用于查找隐藏的对话框"""
    desktop = Desktop(backend="uia")
    for win in desktop.windows():
        try:
            yield win
        except Exception:
            continue


def find_blob_dialog_by_title():
    """使用 Windows API 直接查找 blob URL 窗口"""
    try:
        import ctypes
        from ctypes import wintypes

        # 定义 Windows API
        user32 = ctypes.WinDLL('user32')
        user32.EnumWindows.argtypes = [ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM), wintypes.LPARAM]
        user32.GetWindowTextW.argtypes = [wintypes.HWND, ctypes.c_wchar_p, wintypes.INT]

        results = []

        def enum_callback(hwnd, lParam):
            title = ctypes.create_unicode_buffer(512)
            user32.GetWindowTextW(hwnd, title, 512)
            if title.value.startswith("blob:http"):
                results.append(hwnd)
            return True

        enum_proc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)(enum_callback)
        user32.EnumWindows(enum_proc, 0)

        if results:
            # 使用 pywinauto 包装窗口句柄
            from pywinauto import WindowSpecification
            return WindowSpecification({'handle': results[0]})
    except Exception as e:
        print(f"find_blob_dialog_by_title 失败: {e}")

    return None


def find_app_window(timeout: int):
    end_time = time.time() + timeout

    while time.time() < end_time:
        for win in iter_visible_windows():
            try:
                title = win.window_text().strip()
                if WINDOW_TITLE_HINT in title:
                    return win
            except Exception:
                continue
        time.sleep(1)

    return None


def launch_app() -> None:
    if not os.path.exists(APP_SHORTCUT):
        raise FileNotFoundError(f"找不到快捷方式: {APP_SHORTCUT}")

    print("正在启动顺发打单...")
    os.startfile(APP_SHORTCUT)


def focus_window(win) -> None:
    try:
        win.restore()
    except Exception:
        pass

    try:
        win.set_focus()
    except Exception:
        pass

    time.sleep(1)


def maximize_window(win) -> None:
    try:
        win.maximize()
        print("已将顺发窗口最大化。")
    except Exception:
        print("顺发窗口最大化失败，继续按当前窗口尺寸执行。")

    time.sleep(1)


def get_window_rect(win):
    rect = win.rectangle()
    width = rect.right - rect.left
    height = rect.bottom - rect.top
    if width <= 0 or height <= 0:
        raise RuntimeError("顺发窗口尺寸无效，无法计算点击坐标。")
    return rect.left, rect.top, width, height


def load_points():
    if not POINTS_FILE.exists():
        return None

    with POINTS_FILE.open("r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, dict):
        raise RuntimeError(f"坐标文件格式无效: {POINTS_FILE}")

    return data


def save_points(points) -> None:
    with POINTS_FILE.open("w", encoding="utf-8") as f:
        json.dump(points, f, ensure_ascii=False, indent=2)

    print(f"已保存坐标文件: {POINTS_FILE}")


def record_points(win) -> None:
    focus_window(win)
    maximize_window(win)
    left, top, width, height = get_window_rect(win)

    print("开始录制点击坐标。")
    print("请勿移动顺发窗口位置或尺寸；后续自动点击会基于窗口相对坐标执行。")

    points = {}
    for step in ACTION_STEPS:
        input(f'请把鼠标移动到"{step}"按钮中心，然后按回车记录坐标: ')
        x, y = pyautogui.position()
        rel_x = (x - left) / width
        rel_y = (y - top) / height

        if not (0 <= rel_x <= 1 and 0 <= rel_y <= 1):
            raise RuntimeError(
                f'记录"{step}"失败：鼠标不在顺发窗口范围内。'
            )

        points[step] = {"rel_x": rel_x, "rel_y": rel_y}
        print(
            f'已记录"{step}": abs=({x}, {y}), '
            f"rel=({rel_x:.4f}, {rel_y:.4f})"
        )

    save_points(points)


def click_relative_point(win, step: str, rel_x: float, rel_y: float, rect: tuple = None) -> None:
    if rect is None:
        rect = get_window_rect(win)
    left, top, width, height = rect
    x = int(left + width * rel_x)
    y = int(top + height * rel_y)

    pyautogui.moveTo(x, y, duration=0.2)
    pyautogui.click()
    print(f'已模拟点击"{step}": ({x}, {y})')


def generate_filename() -> str:
    """生成带日期的文件名和完整路径"""
    date_str = datetime.now().strftime("%Y%m%d")
    filename = SAVE_CONFIG["filename_template"].replace("{date}", date_str)

    # 确保目录存在
    save_dir = SAVE_CONFIG["directory"]
    os.makedirs(save_dir, exist_ok=True)

    return os.path.join(save_dir, filename)


def find_save_dialog(timeout: int = None):
    """查找另存为对话框（支持多种可能的标题）"""
    if timeout is None:
        timeout = SAVE_CONFIG["dialog_timeout"]

    end_time = time.time() + timeout
    dialog_titles = ["另存为", "Save As", "保存", "Export", "导出", "Dialog"]

    while time.time() < end_time:
        # 方法1: 使用 iter_all_windows 枚举
        for win in iter_all_windows():
            try:
                title = win.window_text().strip()
                # 检查是否为文件保存对话框
                if any(t in title for t in dialog_titles):
                    return win
                if title.startswith("blob:http"):
                    return win
            except Exception:
                continue

        # 方法2: 使用 Windows API 直接查找 blob 窗口
        blob_win = find_blob_dialog_by_title()
        if blob_win:
            return blob_win

        time.sleep(0.5)

    return None


def handle_save_dialog():
    """处理文件保存对话框 - 输入路径并保存"""
    print("等待文件保存对话框...")
    save_dialog = find_save_dialog()

    if save_dialog is None:
        print("未找到文件保存对话框，所有窗口如下：")
        dump_all_windows()
        raise RuntimeError("未找到文件保存对话框，请检查顺发应用是否正常弹出保存窗口。")

    print(f"已找到保存对话框: {save_dialog.window_text().strip()}")

    # 聚焦到对话框
    try:
        save_dialog.set_focus()
        time.sleep(0.5)
    except Exception as e:
        print(f"聚焦对话框失败: {e}")

    # 生成完整文件路径
    file_path = generate_filename()
    print(f"准备保存到: {file_path}")

    # 方法1: 尝试直接定位文件名输入框
    try:
        # 尝试找到文件名编辑框（常见控件名）
        edit = save_dialog.child_window(auto_id="FileNameControl", control_type="Edit")
        if not edit.exists():
            edit = save_dialog.child_window(title="文件名(&N):", control_type="Edit")
        if edit.exists():
            edit.set_text(file_path)
            print("已通过控件设置文件路径")
        else:
            raise RuntimeError("未找到文件名输入框控件")
    except Exception:
        # 方法2: 使用键盘输入（兼容性更好）
        print("使用键盘输入方式...")
        # 清空已有内容并输入新路径
        send_keys("^a")  # Ctrl+A 全选
        time.sleep(0.2)
        send_keys(file_path)
        time.sleep(0.3)

    # 点击保存按钮或回车
    try:
        # 尝试找到并点击保存按钮
        save_btn = save_dialog.child_window(title="保存(&S)", control_type="Button")
        if save_btn.exists():
            save_btn.click()
            print("已点击保存按钮")
        else:
            # 回车键保存
            send_keys("{ENTER}")
            print("已发送回车键保存")
    except Exception:
        # 回车键作为后备
        send_keys("{ENTER}")
        print("已发送回车键保存")

    # 等待保存完成（检查对话框是否消失）
    time.sleep(0.5)
    elapsed = 0
    timeout = SAVE_CONFIG["save_button_timeout"]
    while elapsed < timeout:
        if not save_dialog.is_visible():
            print("文件保存成功。")
            return True
        time.sleep(0.3)
        elapsed += 0.3

    print("保存命令已发送，请确认文件是否正常保存。")
    return True


def perform_action_steps(win, points) -> None:
    # 开始前只聚焦和最大化一次
    focus_window(win)
    maximize_window(win)
    rect = get_window_rect(win)

    for step in ACTION_STEPS:
        point = points.get(step)
        if not point:
            raise RuntimeError(f'坐标文件缺少"{step}"的点击位置。')

        # 点击"下载文件"前打印当前窗口列表
        if step == "下载文件":
            print("=== 点击下载文件前的当前窗口列表 ===")
            dump_visible_windows()
            print("===================================")

        click_relative_point(
            win,
            step,
            rel_x=float(point["rel_x"]),
            rel_y=float(point["rel_y"]),
            rect=rect,
        )

        # 点击"提交"后等待5秒
        if step == "提交":
            print(f"已点击提交，等待 {SUBMIT_WAIT_SECONDS} 秒...")
            time.sleep(SUBMIT_WAIT_SECONDS)
        # 点击"下载文件"后处理文件保存对话框
        elif step == "下载文件":
            print("已点击下载文件，等待保存对话框...")
            # 等待对话框出现（blob URL 窗口可能需要更长时间）
            time.sleep(3)
            handle_save_dialog()
        else:
            time.sleep(CLICK_INTERVAL_SECONDS)


def dump_visible_windows() -> None:
    print("当前可见顶层窗口:")
    found = False

    for win in iter_visible_windows():
        try:
            title = win.window_text().strip() or "<无标题>"
            class_name = win.element_info.class_name or "<未知类名>"
            print(f"- 标题: {title} | 类名: {class_name}")
            found = True
        except Exception:
            continue

    if not found:
        print("- 未枚举到可见窗口")


def dump_all_windows() -> None:
    """打印所有窗口（包括不可见的），用于调试"""
    print("当前所有窗口（包括不可见的）:")
    found = False

    for win in iter_all_windows():
        try:
            title = win.window_text().strip() or "<无标题>"
            class_name = win.element_info.class_name or "<未知类名>"
            visible = win.is_visible() if hasattr(win, 'is_visible') else True
            visibility = "可见" if visible else "不可见"
            print(f"- 标题: {title} | 类名: {class_name} | 状态: {visibility}")
            found = True
        except Exception:
            continue

    if not found:
        print("- 未枚举到任何窗口")


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--record",
        action="store_true",
        help='录制"全部订单 / 导出 / 提交"三个按钮的鼠标坐标',
    )
    parser.add_argument(
        "--interval-seconds",
        type=int,
        default=0,
        help="按固定秒数循环执行；例如 300 表示每 5 分钟执行一次",
    )
    parser.add_argument(
        "--save-dir",
        type=str,
        default=None,
        help="文件保存目录，覆盖配置文件中的设置",
    )
    parser.add_argument(
        "--filename",
        type=str,
        default=None,
        help="文件名模板，支持 {date} 占位符，如: 订单_{date}.xlsx",
    )
    return parser.parse_args()


def run_once(args) -> int:
    try:
        win = find_app_window(timeout=5)
        if win is None:
            launch_app()
            time.sleep(3)
            win = find_app_window(timeout=40)

        if win is None:
            print("未找到顺发窗口。")
            dump_visible_windows()
            return 2

        title = win.window_text().strip() or "<无标题>"
        print(f"已定位顺发窗口: {title}")
        maximize_window(win)

        if args.record:
            record_points(win)
            return 0

        points = load_points()
        if points is None:
            print("未找到坐标文件，请先执行:")
            print("python sf_open_monthly_settlement_nav.py --record")
            return 2

        print('开始顺序执行"全部订单 -> 导出 -> 提交"...')
        perform_action_steps(win, points)
        print("操作完成。")
        return 0
    except KeyboardInterrupt:
        print("操作已取消。")
        return 1
    except Exception as exc:
        print(f"执行失败: {exc}")
        return 1


def main() -> int:
    if sys.platform != "win32":
        print("此脚本需要在 Windows 环境中运行。")
        return 1

    pyautogui.FAILSAFE = True
    pyautogui.PAUSE = 0.2
    args = parse_args()

    # 应用命令行参数覆盖配置
    if args.save_dir:
        SAVE_CONFIG["directory"] = args.save_dir
    if args.filename:
        SAVE_CONFIG["filename_template"] = args.filename

    # 打印当前保存配置
    print(f"保存目录: {SAVE_CONFIG['directory']}")
    print(f"文件名模板: {SAVE_CONFIG['filename_template']}")

    if args.record and args.interval_seconds > 0:
        print("--record 与 --interval-seconds 不能同时使用。")
        return 1

    if args.interval_seconds <= 0:
        return run_once(args)

    print(
        "已进入循环执行模式，"
        f"每 {args.interval_seconds} 秒执行一次。按 Ctrl+C 可停止。"
    )

    while True:
        started_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"[{started_at}] 开始新一轮执行...")
        run_once(args)
        print(f"等待 {args.interval_seconds} 秒后执行下一轮。")
        time.sleep(args.interval_seconds)


if __name__ == "__main__":
    raise SystemExit(main())
