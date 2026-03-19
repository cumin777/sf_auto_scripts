# 顺发订单筛选脚本

## 用途

从顺发打单导出的Excel订单文件中，筛选出指定收件人地址的订单数据。

**默认筛选条件**：收件人详细地址为 `广东省****3栋9楼`

## 依赖安装

```bash
pip install pandas openpyxl
```

## 使用方法

```bash
# 筛选今天的订单
python sf_filter_orders.py

# 筛选指定日期的订单
python sf_filter_orders.py 20260318
```

## 文件说明

| 脚本 | 说明 |
|------|------|
| `sf_open_monthly_settlement_nav.py` | 顺发打单自动化脚本，导出原始订单Excel |
| `sf_filter_orders.py` | 筛选脚本，从原始Excel中提取指定地址的订单 |

## 输入输出

- **输入文件**：`C:\Orders\Export\顺发订单_YYYYMMDD.xlsx`
- **输出文件**：`C:\Orders\Export\顺发订单_YYYYMMDD_筛选后.xlsx`

原始文件不会被修改。

## 配置

如需修改筛选地址或文件路径，编辑脚本中的以下配置：

```python
# 修改筛选条件
FILTER_ADDRESS = "广东省****3栋9楼"

# 修改文件路径（需与 sf_open_monthly_settlement_nav.py 保持一致）
SAVE_CONFIG = {
    "directory": r"C:\Orders\Export",
    "filename_template": "顺发订单_{date}.xlsx",
}
```

## 运行环境

- 需要在Windows环境运行（文件路径为Windows路径）
- Python 3.7+
