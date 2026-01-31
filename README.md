# 午餐三聯式繳費單產生器

[![.NET](https://img.shields.io/badge/.NET-10.0-512BD4)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE)

## 專案簡介

這是一個 WPF 桌面應用程式，用於從 Access 資料庫讀取學生餐費資料，並產生符合台灣超商代收規範的三聯式繳費單 PDF 報表。

### 主要功能

- ? 讀取 Access `.accdb` 資料庫
- ? 產生 A4 三聯式繳費單（學校存根聯、繳費人存根聯、超商繳款聯）
- ? 支援超商代收條碼（Code39 格式）
- ? 匯出 PDF 報表
- ? 直接列印功能
- ? 預覽功能
- ? 可自訂報表備註

### 技術規格

- **框架**: .NET 10 (x64)
- **UI**: WPF (Windows Presentation Foundation)
- **報表引擎**: RDLC (Report Definition Language Client-side)
- **資料庫**: Microsoft Access (.accdb)
- **條碼庫**: ZXing.Net
- **報表庫**: Microsoft.Reporting.NETCore

## 系統需求

### 開發環境
- Visual Studio 2022 或更新版本
- .NET 10 SDK
- Windows 10/11 (x64)

### 執行環境
- Windows 10/11 (x64)
- .NET 10 Runtime
- Microsoft Access Database Engine 2016 (x64) - [下載連結](https://www.microsoft.com/download/details.aspx?id=54920)

## 快速開始

### 1. 複製專案
```bash
git clone <repository-url>
cd st_lunch_bill_report
```

### 2. 還原套件
```bash
dotnet restore
```

### 3. 建置專案
```bash
dotnet build
```

### 4. 執行程式
```bash
dotnet run
```

## 使用說明

### 基本操作流程

1. **選擇資料庫**
   - 點擊「選擇資料庫...」按鈕
   - 選擇包含學生餐費資料的 `.accdb` 檔案

2. **選擇資料來源**
   - 從下拉選單選擇資料表或查詢
   - 系統會自動載入並驗證資料

3. **設定匯出資料夾**（可選）
   - 點擊「匯出資料夾...」按鈕
   - 選擇 PDF 檔案的輸出位置

4. **自訂報表備註**（可選）
   - 在「報表備註」欄位輸入自訂文字
   - 此文字會顯示在第二聯的備註區域

5. **產生報表**
   - **預覽**: 在 PDF 閱讀器中預覽報表
   - **匯出 PDF**: 將報表儲存為 PDF 檔案
   - **列印**: 直接傳送到印表機

### 資料庫欄位對應

程式會自動對應以下欄位（可透過 `fieldmap.json` 自訂）：

| 標準欄位名稱 | Access 資料庫欄位 | 說明 |
|------------|-----------------|------|
| ClassName | 班級 | 學生班級 |
| ParentName | 姓名 | 學生或家長姓名 |
| StudentId | 學號 | 學生學號 |
| SeatNumber | 座號 | 座號 |
| LunchFee | 學期午餐費用 | 午餐費 |
| FreshMilkFee | 新生訓練餐費 | 新生訓練費 |
| SchoolMealFee | 暑期輔導餐費 | 暑期餐費 |
| AfterSchoolFee | 退前一學期餐費 | 退費金額 |
| AgencyFee | 超商代收手續費 | 手續費 |
| TotalFee | 應繳金額 | 總金額 |
| TotalFeeAlt | 應繳金額 | 替代總金額 |
| FeeDetails | 應繳金額國字 | 中文大寫金額 |
| ParentNote | 繳費期限 | 期限（YYYMMDD格式） |
| ContactInfo | 銷帳編號 | 銷帳編號 |
| Barcode1 | 條碼1序號碼 | 第一段條碼 |
| Barcode2 | 條碼2序號碼 | 第二段條碼 |
| Barcode3 | 條碼3序號碼 | 第三段條碼 |

### 欄位對應設定 (fieldmap.json)

如果您的資料庫欄位名稱不同，可以編輯 `fieldmap.json` 檔案來自訂對應：

```json
{
  "ClassName": "班級",
  "ParentName": "學生姓名",
  "StudentId": "學號",
  ...
}
```

## 報表格式說明

### 三聯式版面配置

```
┌─────────────────────────────────────┐
│  第一聯：學校存根聯 (8.6cm)          │
│  - 標題、學生資訊                   │
│  - 費用明細表                       │
│  - 繳費資訊                         │
│  - 條碼區域（右側）                 │
├─────────────────────────────────────┤ ← 虛線（5mm間距）
│  第二聯：繳費人存根聯 (8.6cm)        │
│  - 學生資訊                         │
│  - 繳費資訊                         │
│  - 三段條碼（橫向排列）             │
│  - 自訂備註                         │
├─────────────────────────────────────┤ ← 虛線（5mm間距）
│  第三聯：超商繳款聯 (9.4cm)          │
│  - 繳費資訊                         │
│  - 三段條碼（橫向排列）             │
│  - 超商標示                         │
└─────────────────────────────────────┘
```

### 頁面設定
- **紙張大小**: A4 (21cm × 29.7cm)
- **邊界**: 上下 1cm、左右 1cm
- **內容寬度**: 19cm
- **各聯間距**: 5mm（含虛線）

## 故障排除

### 常見問題

**Q: 無法開啟資料庫**
- 確認已安裝 Microsoft Access Database Engine 2016 (x64)
- 檢查資料庫檔案是否損壞
- 確認檔案路徑不包含特殊字元

**Q: 資料驗證失敗**
- 檢查必要欄位是否存在：班級、姓名、學號、條碼1-3
- 確認欄位名稱是否符合 `fieldmap.json` 的對應
- 查看 Log 區域的詳細錯誤訊息

**Q: 條碼無法顯示**
- 確認條碼欄位內容為純數字
- 條碼長度應符合超商代收規範
- 檢查 ZXing.Net 套件是否正確安裝

**Q: PDF 無法開啟**
- 確認系統已安裝 PDF 閱讀器（如 Adobe Acrobat Reader）
- 檢查匯出路徑的寫入權限
- 嘗試以系統管理員身分執行程式

## 專案結構

```
st_lunch_bill_report/
├── MainWindow.xaml          # 主視窗 UI
├── MainWindow.xaml.cs       # 主視窗邏輯
├── App.xaml                 # 應用程式設定
├── Services/
│   ├── DatabaseService.cs   # 資料庫存取服務
│   └── ReportService.cs     # 報表產生服務
├── Reports/
│   └── LunchBill3Part.rdlc  # RDLC 報表定義
├── fieldmap.json            # 欄位對應設定
├── README.md                # 本文件
└── spec.md                  # 詳細規格文件
```

## 開發文件

詳細的技術規格和開發指南，請參考：
- [spec.md](spec.md) - 完整系統規格書
- [fieldmap.json](fieldmap.json) - 欄位對應設定範例

## 授權

本專案採用 MIT 授權條款。

## 聯絡資訊

如有問題或建議，請聯絡開發團隊。

---

**最後更新**: 2025-01-31
