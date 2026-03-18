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


APP_SHORTCUT = r"C:\Users\86138\Desktop\顺发打单.lnk"
WINDOW_TITLE_HINT = "顺发"
POINTS_FILE = Path(__file__).with_name("sf_open_monthly_settlement_nav_points.json")
ACTION_STEPS = ["全部订单", "导出", "提交"]
CLICK_INTERVAL_SECONDS = 1.2


def iter_visible_windows():
    desktop = Desktop(backend="uia")
    for win in desktop.windows():
        try:
            if not win.is_visible():
                continue
            yield win
        except Exception:
            continue


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
        input(f"请把鼠标移动到“{step}”按钮中心，然后按回车记录坐标: ")
        x, y = pyautogui.position()
        rel_x = (x - left) / width
        rel_y = (y - top) / height

        if not (0 <= rel_x <= 1 and 0 <= rel_y <= 1):
            raise RuntimeError(
                f"记录“{step}”失败：鼠标不在顺发窗口范围内。"
            )

        points[step] = {"rel_x": rel_x, "rel_y": rel_y}
        print(
            f"已记录“{step}”: abs=({x}, {y}), "
            f"rel=({rel_x:.4f}, {rel_y:.4f})"
        )

    save_points(points)


def click_relative_point(win, step: str, rel_x: float, rel_y: float) -> None:
    focus_window(win)
    maximize_window(win)
    left, top, width, height = get_window_rect(win)
    x = int(left + width * rel_x)
    y = int(top + height * rel_y)

    pyautogui.moveTo(x, y, duration=0.2)
    pyautogui.click()
    print(f"已模拟点击“{step}”: ({x}, {y})")


def perform_action_steps(win, points) -> None:
    for step in ACTION_STEPS:
        point = points.get(step)
        if not point:
            raise RuntimeError(f"坐标文件缺少“{step}”的点击位置。")

        click_relative_point(
            win,
            step,
            rel_x=float(point["rel_x"]),
            rel_y=float(point["rel_y"]),
        )
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


def parse_args():
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--record",
        action="store_true",
        help="录制“全部订单 / 导出 / 提交”三个按钮的鼠标坐标",
    )
    parser.add_argument(
        "--interval-seconds",
        type=int,
        default=0,
        help="按固定秒数循环执行；例如 300 表示每 5 分钟执行一次",
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

        print("开始顺序执行“全部订单 -> 导出 -> 提交”...")
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
