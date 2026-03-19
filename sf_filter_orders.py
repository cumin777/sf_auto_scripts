# -*- coding: utf-8 -*-
"""
筛选顺发订单Excel文件，只保留指定收件人地址的行。

Usage:
    python sf_filter_orders.py
"""

import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import openpyxl

# 文件配置（与 sf_open_monthly_settlement_nav.py 保持一致）
SAVE_CONFIG = {
    "directory": r"C:\Orders\Export",
    "filename_template": "顺发订单_{date}.xlsx",
}

# 筛选条件
FILTER_ADDRESS = "广东省****3栋9楼"


def get_xlsx_path(date_str: str = None) -> str:
    """
    获取xlsx文件路径
    :param date_str: 日期字符串，格式YYYYMMDD。如果为None，使用今天
    :return: xlsx文件的完整路径
    """
    if date_str is None:
        date_str = datetime.now().strftime("%Y%m%d")

    filename = SAVE_CONFIG["filename_template"].replace("{date}", date_str)
    return os.path.join(SAVE_CONFIG["directory"], filename)


def filter_orders(input_path: str, output_path: str = None, address: str = FILTER_ADDRESS) -> pd.DataFrame:
    """
    筛选订单数据
    :param input_path: 输入xlsx文件路径
    :param output_path: 输出xlsx文件路径，如果为None则自动生成
    :param address: 要筛选的收件人地址
    :return: 筛选后的DataFrame
    """
    # 读取Excel文件
    print(f"正在读取文件: {input_path}")
    df = pd.read_excel(input_path)

    print(f"原始数据行数: {len(df)}")
    print(f"列名: {list(df.columns)}")

    # 查找收件人详细地址列（可能的各种列名）
    address_column = None
    possible_names = ["收件人详细地址", "收件人地址", "详细地址", "地址"]

    for col in df.columns:
        if any(name in str(col) for name in possible_names):
            address_column = col
            break

    if address_column is None:
        print("未找到收件人地址列，可用列名如下:")
        for i, col in enumerate(df.columns, 1):
            print(f"  {i}. {col}")
        raise ValueError("无法确定收件人地址列名，请手动检查Excel文件")

    print(f"使用列名: '{address_column}'")

    # 筛选数据 - 支持模糊匹配（regex=False避免特殊字符被当作正则表达式）
    filtered_df = df[df[address_column].astype(str).str.contains(address, na=False, regex=False)]

    print(f"筛选后数据行数: {len(filtered_df)}")

    # 打印筛选结果
    if len(filtered_df) > 0:
        print("\n筛选结果预览:")
        print(filtered_df.head().to_string())
    else:
        print("警告: 没有找到匹配的记录")

    # 保存结果
    if output_path is None:
        base_name = Path(input_path).stem
        output_dir = Path(input_path).parent
        output_path = output_dir / f"{base_name}_筛选后.xlsx"

    filtered_df.to_excel(output_path, index=False, engine='openpyxl')
    print(f"\n已保存筛选结果到: {output_path}")

    return filtered_df


def main():
    if sys.platform != "win32":
        print("注意: 此脚本配置的路径为Windows路径，在非Windows系统上可能无法直接访问。")
        print("如需在其他系统使用，请修改SAVE_CONFIG中的路径。")

    # 获取文件路径
    xlsx_path = get_xlsx_path()

    # 检查文件是否存在
    if not os.path.exists(xlsx_path):
        print(f"文件不存在: {xlsx_path}")
        print("\n你可以:")
        print("1. 指定其他日期: python sf_filter_orders.py 20260318")
        print("2. 修改脚本中的SAVE_CONFIG路径")
        return 1

    # 解析命令行参数（支持指定日期）
    if len(sys.argv) > 1:
        date_arg = sys.argv[1]
        try:
            # 验证日期格式
            datetime.strptime(date_arg, "%Y%m%d")
            xlsx_path = get_xlsx_path(date_arg)
        except ValueError:
            print(f"无效的日期格式: {date_arg}，应为YYYYMMDD格式")
            return 1

    try:
        filter_orders(xlsx_path)
        return 0
    except Exception as e:
        print(f"筛选失败: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
