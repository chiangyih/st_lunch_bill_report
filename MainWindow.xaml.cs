using System.Data;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Microsoft.Win32;
using st_lunch_bill_report.Services;

namespace st_lunch_bill_report;

/// <summary>
/// 午餐三聯式繳費單產生器主視窗
/// </summary>
public partial class MainWindow : Window
{
    private readonly DatabaseService _databaseService;
    private readonly ReportService _reportService;
    private readonly List<string> _tempFiles = [];
    private DataTable? _currentData;
    private string? _currentDatabasePath;
    private string? _currentExportFolder;

    public MainWindow()
    {
        InitializeComponent();
        _databaseService = new DatabaseService();
        _reportService = new ReportService();
        
        // 初始化學年度與學期
        var now = DateTime.Now;
        var currentRocYear = now.Year - 1911;
        
        // 根據月份判斷學年度與學期
        // 2-7月：上學期結束，為當前學年度第2學期
        // 8-1月：下學期，8月開始為新學年度第1學期
        if (now.Month >= 2 && now.Month <= 7)
        {
            // 2-7月：學年度為前一年，第2學期
            currentRocYear--;
            CmbSemester.SelectedIndex = 1; // 第2學期
        }
        else
        {
            // 8-12月或1月：第1學期
            if (now.Month == 1)
            {
                currentRocYear--; // 1月仍屬於前一學年度
            }
            CmbSemester.SelectedIndex = 0; // 第1學期
        }
        
        // 設定學年度下拉選單（選項為 114~124）
        var yearIndex = currentRocYear - 114;
        if (yearIndex >= 0 && yearIndex < CmbSchoolYear.Items.Count)
        {
            CmbSchoolYear.SelectedIndex = yearIndex;
        }
        else
        {
            CmbSchoolYear.SelectedIndex = 0; // 預設選擇第一個（114）
        }
        
        Log("INFO", "應用程式啟動");
    }

    #region 日誌記錄

    private void Log(string level, string message, Exception? ex = null)
    {
        var timestamp = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff");
        var logEntry = $"{timestamp} [{level}] {message}{Environment.NewLine}";
        if (ex != null)
        {
            logEntry += $"Exception Detail:{Environment.NewLine}{ex}{Environment.NewLine}";
        }
        
        Dispatcher.Invoke(() =>
        {
            TxtLog.AppendText(logEntry);
            TxtLog.ScrollToEnd();
        });
    }

    #endregion

    #region UI 狀態管理

    private void UpdateButtonStates()
    {
        var hasDatabase = !string.IsNullOrEmpty(_currentDatabasePath);
        var hasDataSource = CmbDataSource.SelectedItem != null;
        var hasExportFolder = !string.IsNullOrEmpty(_currentExportFolder);
        var hasData = _currentData != null && _currentData.Rows.Count > 0;

        CmbDataSource.IsEnabled = hasDatabase;
        BtnPreview.IsEnabled = hasData;
        BtnExportPdf.IsEnabled = hasData && hasExportFolder;
        BtnPrint.IsEnabled = hasData;

        TxtRecordCount.Text = $"資料筆數：{_currentData?.Rows.Count ?? 0}";
    }

    private void SetStatus(string status)
    {
        TxtStatus.Text = status;
    }

    #endregion

    #region 資料庫檔案操作

    private void BtnSelectDatabase_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFileDialog
        {
            Title = "選擇 Access 資料庫",
            Filter = "Access 資料庫 (*.accdb)|*.accdb",
            CheckFileExists = true
        };

        if (dialog.ShowDialog() == true)
        {
            _currentDatabasePath = dialog.FileName;
            TxtDatabasePath.Text = _currentDatabasePath;
            Log("INFO", $"已選擇資料庫：{_currentDatabasePath}");

            LoadDataSources();
        }
    }

    private void LoadDataSources()
    {
        try
        {
            SetStatus("正在載入資料來源...");
            CmbDataSource.Items.Clear();
            _currentData = null;

            if (string.IsNullOrEmpty(_currentDatabasePath))
                return;

            var sources = _databaseService.GetTableAndQueryNames(_currentDatabasePath);
            foreach (var source in sources)
            {
                CmbDataSource.Items.Add(source);
            }

            Log("INFO", $"已載入 {sources.Count} 個資料來源");
            SetStatus("就緒");
        }
        catch (Exception ex)
        {
            Log("ERROR", $"載入資料來源失敗：{ex.Message}", ex);
            MessageBox.Show($"無法載入資料來源：\n\n{ex.Message}\n\n請確認已安裝 Microsoft Access Database Engine 64-bit",
                "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("錯誤");
        }
        finally
        {
            UpdateButtonStates();
        }
    }

    private void BtnCopyDbPath_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(TxtDatabasePath.Text))
        {
            Clipboard.SetText(TxtDatabasePath.Text);
            Log("INFO", "已複製資料庫路徑到剪貼簿");
        }
    }

    #endregion

    #region 資料來源操作

    private void CmbDataSource_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (CmbDataSource.SelectedItem == null)
            return;

        LoadData();
    }

    private void LoadData()
    {
        try
        {
            var sourceName = CmbDataSource.SelectedItem?.ToString();
            if (string.IsNullOrEmpty(_currentDatabasePath) || string.IsNullOrEmpty(sourceName))
                return;

            SetStatus("正在載入資料...");
            Log("INFO", $"正在從 [{sourceName}] 載入資料...");

            _currentData = _databaseService.LoadData(_currentDatabasePath, sourceName);

            // 驗證資料
            var validationErrors = _databaseService.ValidateData(_currentData);
            if (validationErrors.Count > 0)
            {
                foreach (var error in validationErrors)
                {
                    Log("WARN", error);
                }
                
                var result = MessageBox.Show(
                    $"資料驗證發現 {validationErrors.Count} 個問題：\n\n" +
                    string.Join("\n", validationErrors.Take(10)) +
                    (validationErrors.Count > 10 ? $"\n...還有 {validationErrors.Count - 10} 個問題" : "") +
                    "\n\n是否仍要繼續？",
                    "資料驗證警告",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.No)
                {
                    _currentData = null;
                    UpdateButtonStates();
                    return;
                }
            }

            // 加入使用者輸入的備註欄位
            if (!_currentData.Columns.Contains("ReportNote"))
            {
                _currentData.Columns.Add("ReportNote", typeof(string));
            }
            var reportNote = TxtReportNote.Text;
            foreach (DataRow row in _currentData.Rows)
            {
                row["ReportNote"] = reportNote;
            }

            // 轉換繳費期限格式：從 MMDD 轉換為 YYYMMDD（民國年月日）
            if (_currentData.Columns.Contains("ParentNote"))
            {
                foreach (DataRow row in _currentData.Rows)
                {
                    var originalValue = row["ParentNote"]?.ToString();
                    if (!string.IsNullOrEmpty(originalValue))
                    {
                        // 若為日期類型，轉換為民國年格式
                        if (row["ParentNote"] is DateTime dateValue)
                        {
                            var rocYear = dateValue.Year - 1911;
                            row["ParentNote"] = $"{rocYear:000}{dateValue:MMdd}";
                        }
                        // 若為 MMDD 格式（4位數字），加上當前民國年
                        else if (originalValue.Length == 4 && int.TryParse(originalValue, out _))
                        {
                            var rocYear = DateTime.Now.Year - 1911;
                            row["ParentNote"] = $"{rocYear:000}{originalValue}";
                        }
                    }
                }
            }

            // 生成條碼圖片
            _reportService.GenerateBarcodeImages(_currentData);

            Log("INFO", $"已載入 {_currentData.Rows.Count} 筆資料");
            SetStatus("就緒");
        }
        catch (Exception ex)
        {
            Log("ERROR", $"載入資料失敗：{ex.Message}", ex);
            MessageBox.Show($"載入資料失敗：\n\n{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            _currentData = null;
            SetStatus("錯誤");
        }
        finally
        {
            UpdateButtonStates();
        }
    }

    #endregion

    #region 匯出設定操作

    private void BtnSelectExportFolder_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new OpenFolderDialog
        {
            Title = "選擇匯出資料夾"
        };

        if (dialog.ShowDialog() == true)
        {
            _currentExportFolder = dialog.FolderName;
            TxtExportFolder.Text = _currentExportFolder;
            Log("INFO", $"已選擇匯出資料夾：{_currentExportFolder}");
            UpdateButtonStates();
        }
    }

    private void BtnCopyExportPath_Click(object sender, RoutedEventArgs e)
    {
        if (!string.IsNullOrEmpty(TxtExportFolder.Text))
        {
            Clipboard.SetText(TxtExportFolder.Text);
            Log("INFO", "已複製匯出路徑到剪貼簿");
        }
    }

    #endregion

    #region 報表操作

    private void BtnPreview_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_currentData == null)
            {
                MessageBox.Show("請先載入資料", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SetStatus("正在產生預覽...");
            Log("INFO", "開始產生預覽 PDF...");

            var tempPath = Path.Combine(Path.GetTempPath(), $"LunchBill_Preview_{DateTime.Now:yyyyMMdd_HHmmss}.pdf");
            _reportService.ExportToPdf(_currentData, tempPath);
            _tempFiles.Add(tempPath);

            Log("INFO", $"預覽 PDF 已產生：{tempPath}");

            // 使用系統預設 PDF 閱讀器開啟
            var psi = new ProcessStartInfo
            {
                FileName = tempPath,
                UseShellExecute = true
            };
            
            var process = Process.Start(psi);
            if (process == null)
            {
                throw new InvalidOperationException("無法開啟 PDF 預覽檔案。請確認已安裝 PDF 閱讀器。");
            }

            Log("INFO", "已開啟預覽");
            SetStatus("就緒");
        }
        catch (Exception ex)
        {
            Log("ERROR", $"預覽失敗：{ex.Message}", ex);
            MessageBox.Show($"預覽失敗：\n\n{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("錯誤");
        }
    }

    private void BtnExportPdf_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_currentData == null)
            {
                MessageBox.Show("請先載入資料", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            if (string.IsNullOrEmpty(_currentExportFolder))
            {
                MessageBox.Show("請先選擇匯出資料夾", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SetStatus("正在匯出 PDF...");
            Log("INFO", "開始匯出 PDF...");

            var fileName = $"LunchBill_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
            var filePath = Path.Combine(_currentExportFolder, fileName);

            _reportService.ExportToPdf(_currentData, filePath);

            Log("INFO", $"PDF 已匯出：{filePath}");
            MessageBox.Show($"PDF 已成功匯出至：\n\n{filePath}", "匯出成功", MessageBoxButton.OK, MessageBoxImage.Information);
            SetStatus("就緒");
        }
        catch (Exception ex)
        {
            Log("ERROR", $"匯出 PDF 失敗：{ex.Message}", ex);
            MessageBox.Show($"匯出 PDF 失敗：\n\n{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("錯誤");
        }
    }

    private void BtnPrint_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (_currentData == null)
            {
                MessageBox.Show("請先載入資料", "提示", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            SetStatus("正在準備列印...");
            Log("INFO", "開始列印...");

            _reportService.Print(_currentData);

            Log("INFO", "列印工作已送出");
            SetStatus("就緒");
        }
        catch (Exception ex)
        {
            Log("ERROR", $"列印失敗：{ex.Message}", ex);
            MessageBox.Show($"列印失敗：\n\n{ex.Message}", "錯誤", MessageBoxButton.OK, MessageBoxImage.Error);
            SetStatus("錯誤");
        }
    }

    #endregion

    #region 程式結束

    private void BtnExit_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
    {
        Log("INFO", "正在關閉程式...");

        // 清理暫存檔案
        foreach (var tempFile in _tempFiles)
        {
            try
            {
                if (File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                    Log("INFO", $"已刪除暫存檔：{tempFile}");
                }
            }
            catch
            {
                // 忽略刪除失敗
            }
        }

        // 釋放資源
        _databaseService.Dispose();
        _reportService.Dispose();

        Log("INFO", "Application shutdown.");
    }

    #endregion
}