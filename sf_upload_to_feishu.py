# -*- coding: utf-8 -*-
"""
将筛选后的顺发订单Excel文件上传到飞书多维表格。

Usage:
    1. 配置下方飞书相关参数
    2. python sf_upload_to_feishu.py
"""

import os
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
import requests

# ==================== 配置区域 ====================

# 飞书应用配置（需要在飞书开放平台获取）
FEISHU_APP_ID = "cli_a9311b297a779cb5"               # 飞书自建应用的App ID
FEISHU_APP_SECRET = "BMOnNe3dbDp60fHxLUAa1gMQjCCbUb0j"       # 飞书自建应用的App Secret
FEISHU_APP_TOKEN = "YyQMbi1nKaWYEmsGMthcHpCanAh"  # 飞书多维表格的app_token
FEISHU_TABLE_ID = "tbl7ITXSKryUFwSA"           # 飞书多维表格的table_id


# Excel文件配置（与 sf_filter_orders.py 保持一致）
EXCEL_CONFIG = {
    "directory": r"C:\Orders\Export",
    "filename_template": "顺发订单_{date}_筛选后.xlsx",
}

# 字段映射：Excel列名 -> 飞书表格field_name（使用字段名，不是field_id）
FIELD_MAPPING = {
    "运单号": "运单号",
    "创建时间": "订单时间",
    "寄件人详细地址": "寄件地址",
    "收件人详细地址": "收件地址",
    "订单状态": "订单状态",
    "收件人姓名": "收件人",
    "收件人手机": "收件人电话",
    "物流产品": "物流产品",
}

# ==================================================


def get_excel_path(date_str: str = None) -> str:
    """获取筛选后的Excel文件路径"""
    # 临时固定文件，用于测试
    return r"C:\Orders\Export\顺发订单_20260319_筛选后.xlsx"

    # 原逻辑（测试后恢复）
    # if date_str is None:
    #     date_str = datetime.now().strftime("%Y%m%d")
    # filename = EXCEL_CONFIG["filename_template"].replace("{date}", date_str)
    # return os.path.join(EXCEL_CONFIG["directory"], filename)


def get_feishu_access_token():
    """获取飞书 tenant_access_token"""
    if FEISHU_APP_ID == "your_app_id" or FEISHU_APP_SECRET == "your_app_secret":
        raise ValueError("请先配置 FEISHU_APP_ID 和 FEISHU_APP_SECRET")

    # 打印配置信息（隐藏部分敏感信息）
    print(f"App ID: {FEISHU_APP_ID}")
    print(f"App Secret: {FEISHU_APP_SECRET[:8]}...{FEISHU_APP_SECRET[-4:]}")

    url = "https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal/"
    headers = {"Content-Type": "application/json; charset=utf-8"}
    data = {"app_id": FEISHU_APP_ID, "app_secret": FEISHU_APP_SECRET}

    try:
        response = requests.post(url, headers=headers, json=data)
        print(f"HTTP状态码test: {response.status_code}")
        result = response.json()
        print(f"API响应: code={result.get('code')}, msg={result.get('msg')}")

        response.raise_for_status()
        if result.get("code") != 0:
            raise RuntimeError(f"获取token失败: {result.get('msg')}")
        return result.get("tenant_access_token")
    except Exception as e:
        print(f"获取飞书access_token失败: {e}")
        return None


def read_excel_data(file_path: str) -> list:
    """读取Excel文件并转换为字典列表"""
    print(f"正在读取文件: {file_path}")
    df = pd.read_excel(file_path)

    if len(df) == 0:
        print("文件中没有数据")
        return []

    print(f"读取到 {len(df)} 条记录")
    return df.to_dict(orient="records")


def build_feishu_record(row: dict) -> dict:
    """将Excel行数据转换为飞书记录格式"""
    fields = {}

    for excel_col, feishu_field in FIELD_MAPPING.items():
        value = row.get(excel_col, "")
        # 跳过空值和NaN
        if value is None or (isinstance(value, float) and pd.isna(value)):
            continue
        if str(value).strip() == "":
            continue
        fields[feishu_field] = str(value).strip()

    return {"fields": fields}


def write_to_feishu(records: list, access_token: str) -> bool:
    """将数据写入飞书多维表格"""
    if not records:
        print("没有数据需要写入")
        return True

    if FEISHU_APP_TOKEN == "your_app_token" or FEISHU_TABLE_ID == "your_table_id":
        raise ValueError("请先配置 FEISHU_APP_TOKEN 和 FEISHU_TABLE_ID")

    # 使用批量创建接口
    url = f"https://open.feishu.cn/open-apis/bitable/v1/apps/{FEISHU_APP_TOKEN}/tables/{FEISHU_TABLE_ID}/records/batch_create"
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": "application/json; charset=utf-8"
    }

    # 构造记录
    feishu_records = [build_feishu_record(record) for record in records]

    # 批量写入所有记录
    batch_size = 500  # 飞书API单次最多500条
    success_count = 0

    for i in range(0, len(feishu_records), batch_size):
        batch = feishu_records[i:i + batch_size]
        body = {"records": batch}

        try:
            # 打印请求体用于调试
            import json
            print(f"发送的请求体: {json.dumps(body, ensure_ascii=False, indent=2)}")

            response = requests.post(url, headers=headers, json=body)
            result = response.json()

            # 打印详细响应用于调试
            print(f"HTTP状态码test: {response.status_code}")
            print(f"API响应: code={result.get('code')}, msg={result.get('msg')}")

            if response.status_code == 400 or result.get("code") != 0:
                # 打印第一条记录的结构帮助调试
                print(f"发送的数据结构: {batch[0]}")
                # 打印完整的API错误响应
                print(f"完整错误响应: {result}")
                # 检查字段错误
                if "data" in result and isinstance(result["data"], dict):
                    field_errors = result["data"].get("fieldErrors", {})
                    if field_errors:
                        for field_id, error_msg in field_errors.items():
                            print(f"字段错误 [{field_id}]: {error_msg}")

            response.raise_for_status()

            if result.get("code") == 0:
                batch_success = len(batch)
                success_count += batch_success
                print(f"批次 {i//batch_size + 1}: 成功写入 {batch_success} 条记录")
            else:
                print(f"批次 {i//batch_size + 1} 写入失败: {result.get('msg')}")
                return False

        except Exception as e:
            print(f"写入飞书多维表格失败: {e}")
            return False

    print(f"总计成功写入 {success_count} 条记录到飞书多维表格")
    return True


def print_config_guide():
    """打印配置指南"""
    print("\n" + "="*50)
    print("飞书配置指南")
    print("="*50)
    print("""
1. 获取 App ID 和 App Secret:
   - 访问飞书开放平台: https://open.feishu.cn/
   - 创建企业自建应用
   - 在「凭证与基础信息」页面获取

2. 获取 App Token 和 Table ID:
   - 打开目标多维表格
   - URL格式: https://example.feishu.cn/base/xxxxx/yyyyy
   - App Token = xxxxx
   - Table ID = yyyyy (点击表格后可在URL中找到)

3. 配置权限:
   - 在飞书开放平台为应用添加权限:
     - bitable:app (查看和上传多维表格)
   - 发布应用并在飞书管理后台启用

4. 配置字段映射:
   - 修改脚本中的 FIELD_MAPPING
   - 左边是Excel列名，右边是飞书表格字段名
""")


def main():
    if sys.platform != "win32":
        print("注意: 此脚本配置的路径为Windows路径")

    print(f"开始上传订单到飞书: {datetime.now()}")

    # 检查配置
    if FEISHU_APP_ID == "your_app_id":
        print_config_guide()
        return 1

    # 获取Excel文件路径
    excel_path = get_excel_path()

    # 支持命令行指定日期
    if len(sys.argv) > 1:
        date_arg = sys.argv[1]
        try:
            datetime.strptime(date_arg, "%Y%m%d")
            excel_path = get_excel_path(date_arg)
        except ValueError:
            print(f"无效的日期格式: {date_arg}，应为YYYYMMDD格式")
            return 1

    # 检查文件是否存在
    if not os.path.exists(excel_path):
        print(f"文件不存在: {excel_path}")
        print("请先运行 sf_filter_orders.py 生成筛选后的文件")
        return 1

    try:
        # 读取Excel数据
        records = read_excel_data(excel_path)
        if not records:
            return 0

        # 获取飞书token
        access_token = get_feishu_access_token()
        if not access_token:
            return 1

        # 写入飞书
        if write_to_feishu(records, access_token):
            print("上传完成")
            return 0
        else:
            return 1

    except Exception as e:
        print(f"上传失败: {e}")
        import traceback
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
