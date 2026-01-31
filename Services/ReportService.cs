using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using Microsoft.Reporting.NETCore;
using ZXing;
using ZXing.Common;
using ZXing.Rendering;

namespace st_lunch_bill_report.Services;

/// <summary>
/// 報表服務 - 處理 RDLC 報表產生與輸出
/// </summary>
public class ReportService : IDisposable
{
    private LocalReport? _report;
    private bool _disposed;

    /// <summary>
    /// 條碼圖片欄位名稱
    /// </summary>
    private static readonly string[] BarcodeImageFields = ["Barcode1_Img", "Barcode2_Img", "Barcode3_Img"];
    private static readonly string[] BarcodeSourceFields = ["Barcode1", "Barcode2", "Barcode3"];

    /// <summary>
    /// 產生條碼圖片並加入 DataTable
    /// </summary>
    public void GenerateBarcodeImages(DataTable data)
    {
        // 新增條碼圖片欄位
        foreach (var field in BarcodeImageFields)
        {
            if (!data.Columns.Contains(field))
            {
                data.Columns.Add(field, typeof(byte[]));
            }
        }

        // 為每筆資料產生條碼圖片
        foreach (DataRow row in data.Rows)
        {
            for (int i = 0; i < BarcodeSourceFields.Length; i++)
            {
                var barcodeValue = row[BarcodeSourceFields[i]]?.ToString();
                if (!string.IsNullOrEmpty(barcodeValue))
                {
                    row[BarcodeImageFields[i]] = GenerateCode39Barcode(barcodeValue);
                }
            }
        }
    }

    /// <summary>
    /// 產生 Code39 條碼圖片
    /// </summary>
    private static byte[] GenerateCode39Barcode(string content)
    {
        // ZXing.Net 會自動加上起訖 * 字元
        var writer = new BarcodeWriterPixelData
        {
            Format = BarcodeFormat.CODE_39,
            Options = new EncodingOptions
            {
                Height = 50,           // 條碼高度（像素）
                Width = 300,           // 條碼寬度（像素）
                Margin = 10,           // Quiet Zone
                PureBarcode = false    // 顯示人類可讀碼
            }
        };

        var pixelData = writer.Write(content);

        // 轉換為 PNG 圖片
        using var bitmap = new Bitmap(pixelData.Width, pixelData.Height, PixelFormat.Format32bppRgb);
        using var ms = new MemoryStream();
        
        var bitmapData = bitmap.LockBits(
            new Rectangle(0, 0, pixelData.Width, pixelData.Height),
            ImageLockMode.WriteOnly,
            PixelFormat.Format32bppRgb);

        try
        {
            System.Runtime.InteropServices.Marshal.Copy(
                pixelData.Pixels, 0, bitmapData.Scan0, pixelData.Pixels.Length);
        }
        finally
        {
            bitmap.UnlockBits(bitmapData);
        }

        bitmap.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    /// <summary>
    /// 初始化報表
    /// </summary>
    private void InitializeReport(DataTable data)
    {
        _report?.Dispose();
        _report = new LocalReport();

        var rdlcPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Reports", "LunchBill3Part.rdlc");
        
        if (!File.Exists(rdlcPath))
        {
            throw new FileNotFoundException($"找不到報表檔案：{rdlcPath}");
        }

        using var stream = new FileStream(rdlcPath, FileMode.Open, FileAccess.Read);
        _report.LoadReportDefinition(stream);
        _report.DataSources.Clear();
        _report.DataSources.Add(new ReportDataSource("ReportData", data));
    }

    /// <summary>
    /// 匯出為 PDF
    /// </summary>
    public void ExportToPdf(DataTable data, string outputPath)
    {
        InitializeReport(data);

        var bytes = _report!.Render("PDF");
        File.WriteAllBytes(outputPath, bytes);
    }

    /// <summary>
    /// 列印報表（透過 PDF 中介檔案）
    /// </summary>
    public void Print(DataTable data)
    {
        // 產生暫存 PDF
        var tempPath = Path.Combine(Path.GetTempPath(), $"LunchBill_Print_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
        ExportToPdf(data, tempPath);

        // 使用系統預設 PDF 閱讀器列印
        var psi = new System.Diagnostics.ProcessStartInfo
        {
            FileName = tempPath,
            Verb = "print",
            UseShellExecute = true
        };

        try
        {
            var process = System.Diagnostics.Process.Start(psi);
            if (process == null)
            {
                throw new InvalidOperationException("無法啟動列印程序。請確認已安裝 PDF 閱讀器。");
            }
        }
        catch
        {
            // 若列印動詞不支援，改用開啟方式
            psi.Verb = "open";
            var process = System.Diagnostics.Process.Start(psi);
            if (process == null)
            {
                throw new InvalidOperationException("無法開啟 PDF 檔案。請確認已安裝 PDF 閱讀器。");
            }
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _report?.Dispose();
        _disposed = true;

        GC.SuppressFinalize(this);
    }
}
