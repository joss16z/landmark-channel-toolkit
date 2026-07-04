# Landmark Channel Toolkit V1

第一版功能：**PPT核对Python / Floorplate Update Report**

上传 Floorplate PPT + Price List Excel，自动生成 Excel 检查报告：

- Summary
- Price Changed
- Missing in PPT
- Extra in PPT
- Area Changed
- Lot Changed
- Duplicates in PPT
- Raw PPT Labels
- Raw Excel Available

## GitHub + Streamlit Cloud 部署

1. 把本文件夹内的所有文件上传到 GitHub repository 根目录：
   - `app.py`
   - `requirements.txt`
   - `README.md`
2. 打开 Streamlit Community Cloud
3. Create app / New app
4. Main file path 填：

```text
app.py
```

5. Deploy

## 使用方法

1. 打开网页
2. 上传 Floorplate PPT/PPTX
3. 上传 Price List Excel（xls/xlsx）
4. 点击 **Generate Floorplate Update Report**
5. 下载生成的 `.xlsx` report

## Excel 识别规则

程序会自动识别以下列名：

- Unit：`Apt #`, `Unit`, `Unit Number`, `Apartment`, `Apt`
- Lot：`Lot #`, `Lot`, `Lot Number`
- Price：`Contract Price`, `Price`, `List Price`
- Internal：`Internal`, `Internal (sqm)`
- External：`External`, `External (sqm)`
- Status：`Status`

如果有 Status 列，只检查 Status 包含 `Available` 的户型；如果没有 Status 列，默认检查所有 Excel 户型。

## PPT 识别规则

支持常见格式：

```text
Unit 401
80 + 11
$1,300,000
```

```text
80 + 10
Lot 50_Unit B202 - $980,000
```

```text
Unit B1001
75 + 10
$1,015,000
```

## 注意

这是 V1，核心是 QA 检查，不会修改 PPT。后续可加入 Auto Fix / Floorplate Generator。
