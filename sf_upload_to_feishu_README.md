# 顺发订单上传到飞书脚本

## 功能说明

将筛选后的顺发订单Excel文件自动上传到飞书多维表格。

## 依赖安装

```bash
pip install pandas openpyxl requests
```

## 配置步骤

### 1. 创建飞书应用

1. 访问飞书开放平台：https://open.feishu.cn/
2. 创建企业自建应用
3. 在「凭证与基础信息」获取 App ID 和 App Secret

### 2. 添加权限

在飞书开放平台 → 权限管理，添加以下权限：

| 权限名称 | 权限标识 |
|----------|----------|
| 查看、评论和创建多维表格 | `bitable:app` |

### 3. 发布应用

1. 版本管理与发布 → 创建新版本
2. 申请发布
3. 在飞书管理后台启用应用

### 4. 配置脚本

编辑 `sf_upload_to_feishu.py`，修改以下配置：

```python
# 飞书应用配置
FEISHU_APP_ID = "cli_a9311b297a779cb5"
FEISHU_APP_SECRET = "BMOnNe3dbDp60fHxLUAa1gMQjCCbUb0j"

# 多维表格配置（从URL中提取）
FEISHU_APP_TOKEN = "YyQMbi1nKaWYEmsGMthcHpCanAh"
FEISHU_TABLE_ID = "tbl7ITXSKryUFwSA"

# Excel文件路径
EXCEL_CONFIG = {
    "directory": r"C:\Orders\Export",
    "filename_template": "顺发订单_{date}_筛选后.xlsx",
}

# 字段映射（Excel列名 -> 飞书表格字段名）
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
```

### 5. 获取 App Token 和 Table ID

打开你的飞书多维表格，URL格式：

```
https://xxx.feishu.cn/base/YyQMbi1nKaWYEmsGMthcHpCanAh?table=tbl7ITXSKryUFwSA
```

- **App Token**: `/base/` 后面的部分 → `YyQMbi1nKaWYEmsGMthcHpCanAh`
- **Table ID**: `table=` 后面的部分 → `tbl7ITXSKryUFwSA`

### 6. 添加应用到多维表格

将应用添加为多维表格的协作者：
1. 打开多维表格
2. 右上角「分享」或「协作」
3. 添加你的应用为「可编辑」权限

## 使用方法

```bash
# 上传今天的筛选后文件
python sf_upload_to_feishu.py

# 上传指定日期的文件
python sf_upload_to_feishu.py 20260318
```

## 完整工作流

```bash
# 1. 导出并筛选订单
python sf_filter_orders.py

# 2. 上传到飞书
python sf_upload_to_feishu.py
```

## 获取表格字段信息

如果字段映射有误，运行此脚本获取正确的字段信息：

```bash
python sf_get_feishu_fields.py
```

## 常见问题

### 403 Forbidden

**原因**：权限不足

**解决**：
1. 确认应用已添加 `bitable:app` 权限
2. 确认应用已发布并启用
3. 确认应用已添加为多维表格协作者

### FieldNameNotFound

**原因**：字段名不匹配

**解决**：运行 `sf_get_feishu_fields.py` 获取正确的字段名，更新 `FIELD_MAPPING`

### app secret invalid

**原因**：App Secret 错误

**解决**：去飞书开放平台重新复制 App Secret

## 相关文档

| 文档 | 链接 |
|------|------|
| 批量创建记录API | https://open.feishu.cn/document/server-docs/docs/bitable-v1/app-table-record/batch_create |
| 多维表格API总览 | https://open.feishu.cn/document/server-docs/docs/bitable-v1/bitable-overview |
| 权限列表 | https://open.feishu.cn/document/server-docs/application-scope/scope-list |

## 文件说明

| 文件 | 说明 |
|------|------|
| `sf_open_monthly_settlement_nav.py` | 顺发打单自动化，导出原始订单Excel |
| `sf_filter_orders.py` | 筛选订单，按地址过滤 |
| `sf_upload_to_feishu.py` | 上传筛选后的订单到飞书 |
| `sf_get_feishu_fields.py` | 获取飞书表格字段信息 |
