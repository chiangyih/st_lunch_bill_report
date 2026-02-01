# 國立新化高工午餐三聯式繳費單產生器

[![.NET](https://img.shields.io/badge/.NET-10.0-512BD4)](https://dotnet.microsoft.com/)
[![Platform](https://img.shields.io/badge/Platform-Windows-blue)](https://www.microsoft.com/windows)
[![License](https://img.shields.io/badge/License-MIT-green)](LICENSE.txt)

## 專案概述
本專案為 Windows 桌面應用程式（WPF / .NET 10），讀取既有 Access `.accdb` 資料庫中的完整報表資料，產生 A4 一人一頁的三聯式繳費單（RDLC），支援超商代收條碼（Code39）與 PDF 輸出。

## 主要功能
- 讀取 Access `.accdb` 資料庫（資料表或查詢）
- 依 `fieldmap.json` 進行欄位名稱映射
- 產生三聯式繳費單 RDLC 報表
- 產生 Code39 條碼圖片（ZXing.Net）
- 預覽、匯出 PDF、列印
- 自訂學年度/學期與報表備註
- 完整 Log 與錯誤處理

## 技術與套件
- **目標框架**：.NET 10 (x64)
- **UI**：WPF
- **報表**：RDLC（Microsoft.Reporting.NETCore）
- **資料庫**：Access (`.accdb`)
- **條碼**：ZXing.Net

## 專案結構
```
st_lunch_bill_report/
├── App.xaml
├── MainWindow.xaml
├── MainWindow.xaml.cs
├── Services/
│   ├── DatabaseService.cs
│   └── ReportService.cs
├── Models/
│   └── PaymentSlipData.cs
├── Reports/
│   └── LunchBill3Part.rdlc
├── fieldmap.json
├── spec.md
└── st_lunch_bill_report.csproj
```

## 環境需求
- Windows 10/11 (x64)
- .NET 10 Desktop Runtime
- Microsoft Access Database Engine 2016 (x64)
- PDF 閱讀器（如 Adobe Acrobat Reader）

## 快速開始
1. 以 Visual Studio 2022 開啟 `st_lunch_bill_report.csproj`
2. 還原套件並建置專案
3. 執行程式後，選擇 `.accdb` 資料庫並選擇資料來源
4. 預覽、匯出 PDF 或列印

## 欄位映射（fieldmap.json）
若 Access 欄位名與報表標準欄位名不同，可透過 `fieldmap.json` 調整映射：
```json
{
  "ClassName": "班級",
  "ParentName": "姓名",
  "StudentId": "學號",
  "SeatNumber": "座號",
  "LunchFee": "學期午餐費用",
  "Barcode1": "條碼1序號碼",
  "Barcode2": "條碼2序號碼",
  "Barcode3": "條碼3序號碼"
}
```

## 版本更新紀錄
| 版本 | 日期 | 變更摘要 |
|------|------|----------|
| v2.0 | 2025-01-31 | 更新為 .NET 10、加入學年度/學期與備註功能、更新 UI 與報表規範、補充條碼與驗證流程 |
| v1.6 | 2024-01-31 | SSD 優化：修正章節編號、新增需求編號、補充修訂歷史 |
| v1.5 | 2024-01-31 | 新增動態 SQL 組建規則與預覽模式規範 |
| v1.0 | 2024-01 | 初版建立 |

## 相關文件
- `spec.md`：完整 SSD 規格書
- `Reports/LunchBill3Part.rdlc`：報表版型
- `fieldmap.json`：欄位映射設定

## 授權
本專案採用 AGPL 3.0，詳見 `LICENSE.txt`。
