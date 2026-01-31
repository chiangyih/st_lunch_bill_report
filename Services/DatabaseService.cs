using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Text.Json;
using st_lunch_bill_report.Models;

namespace st_lunch_bill_report.Services;

/// <summary>
/// 資料庫服務 - 處理 Access .accdb 資料庫操作
/// </summary>
public class DatabaseService : IDisposable
{
    private OleDbConnection? _connection;
    private Dictionary<string, string>? _fieldMapping;
    private bool _disposed;

    /// <summary>
    /// 標準欄位名稱（RDLC 使用）- 使用英文以符合 CLS 標準
    /// </summary>
    private static readonly string[] CanonicalFields =
    [
        "ClassName", "ParentName", "StudentId", "SeatNumber",
        "LunchFee", "FreshMilkFee", "SchoolMealFee", "AfterSchoolFee",
        "AgencyFee", "TotalFee", "TotalFeeAlt", "FeeDetails",
        "ParentNote", "ContactInfo", "Barcode1", "Barcode2", "Barcode3"
    ];

    /// <summary>
    /// 預設欄位映射（標準欄位名 -> Access 欄位名）
    /// 已對照 test.accdb 實際欄位名稱
    /// </summary>
    private static readonly Dictionary<string, string> DefaultFieldMapping = new()
    {
        ["ClassName"] = "班級",
        ["ParentName"] = "姓名",
        ["StudentId"] = "學號",
        ["SeatNumber"] = "座號",
        ["LunchFee"] = "學期午餐金額",
        ["FreshMilkFee"] = "新生訓練午餐費",
        ["SchoolMealFee"] = "暑期輔導午餐費",
        ["AfterSchoolFee"] = "退前一學期午餐費",
        ["AgencyFee"] = "超商代收手續費",
        ["TotalFee"] = "應繳金額",
        ["TotalFeeAlt"] = "實繳金額",
        ["FeeDetails"] = "實繳金額中文",
        ["ParentNote"] = "繳費截止日",
        ["ContactInfo"] = "原始銷帳編號",
        ["Barcode1"] = "第1段條碼",
        ["Barcode2"] = "第2段條碼",
        ["Barcode3"] = "第3段條碼"
    };

    public DatabaseService()
    {
        LoadFieldMapping();
    }

    /// <summary>
    /// 載入欄位映射設定
    /// </summary>
    private void LoadFieldMapping()
    {
        var configPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "fieldmap.json");
        
        if (File.Exists(configPath))
        {
            try
            {
                // 使用 UTF-8 編碼讀取
                var json = File.ReadAllText(configPath, System.Text.Encoding.UTF8);
                _fieldMapping = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                
                // 驗證映射是否有效
                if (_fieldMapping != null && _fieldMapping.Count > 0)
                {
                    return;
                }
            }
            catch
            {
                // 讀取失敗，使用預設映射
            }
        }

        // 使用預設映射
        _fieldMapping = new Dictionary<string, string>(DefaultFieldMapping);
    }

    /// <summary>
    /// 取得連線字串
    /// </summary>
    private static string GetConnectionString(string databasePath)
    {
        return $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={databasePath};Persist Security Info=False;";
    }

    /// <summary>
    /// 取得資料表和查詢名稱
    /// </summary>
    public List<string> GetTableAndQueryNames(string databasePath)
    {
        var result = new List<string>();

        using var connection = new OleDbConnection(GetConnectionString(databasePath));
        connection.Open();

        // 取得使用者資料表
        var tables = connection.GetSchema("Tables");
        foreach (DataRow row in tables.Rows)
        {
            var tableType = row["TABLE_TYPE"]?.ToString();
            var tableName = row["TABLE_NAME"]?.ToString();

            if (!string.IsNullOrEmpty(tableName) && 
                (tableType == "TABLE" || tableType == "VIEW") &&
                !tableName.StartsWith("MSys"))
            {
                result.Add(tableName);
            }
        }

        return result.Order().ToList();
    }

    /// <summary>
    /// 取得資料表的實際欄位名稱
    /// </summary>
    private List<string> GetActualColumnNames(string sourceName)
    {
        var columns = new List<string>();
        
        using var command = new OleDbCommand($"SELECT TOP 1 * FROM [{sourceName}]", _connection);
        using var reader = command.ExecuteReader(CommandBehavior.SchemaOnly);
        var schemaTable = reader.GetSchemaTable();
        
        if (schemaTable != null)
        {
            foreach (DataRow row in schemaTable.Rows)
            {
                var columnName = row["ColumnName"]?.ToString();
                if (!string.IsNullOrEmpty(columnName))
                {
                    columns.Add(columnName);
                }
            }
        }

        return columns;
    }

    /// <summary>
    /// 載入資料並套用欄位映射
    /// </summary>
    public DataTable LoadData(string databasePath, string sourceName)
    {
        _connection?.Close();
        _connection = new OleDbConnection(GetConnectionString(databasePath));
        _connection.Open();

        // 先取得實際欄位名稱
        var actualColumns = GetActualColumnNames(sourceName);

        // 動態組建 SQL 語法（只選擇存在的欄位）
        var sql = BuildSelectSql(sourceName, actualColumns);

        using var command = new OleDbCommand(sql, _connection);
        using var adapter = new OleDbDataAdapter(command);
        
        var dataTable = new DataTable("ReportData");
        adapter.Fill(dataTable);

        return dataTable;
    }

    /// <summary>
    /// 依據映射設定動態組建 SELECT SQL 語法
    /// </summary>
    private string BuildSelectSql(string sourceName, List<string> actualColumns)
    {
        var selectClauses = new List<string>();
        var orderByFields = new List<string>();

        foreach (var canonicalField in CanonicalFields)
        {
            // 取得映射的 Access 欄位名
            var accessField = _fieldMapping!.TryGetValue(canonicalField, out var mapped) 
                ? mapped 
                : canonicalField;

            // 檢查欄位是否存在於實際資料表中
            var matchedColumn = actualColumns.FirstOrDefault(c => 
                c.Equals(accessField, StringComparison.OrdinalIgnoreCase));

            if (matchedColumn != null)
            {
                // 欄位存在
                if (canonicalField.Equals(matchedColumn, StringComparison.OrdinalIgnoreCase))
                {
                    selectClauses.Add($"[{matchedColumn}]");
                }
                else
                {
                    selectClauses.Add($"[{matchedColumn}] AS [{canonicalField}]");
                }

                // 收集排序欄位
                if (canonicalField == "ClassName" || canonicalField == "SeatNumber" || canonicalField == "StudentId")
                {
                    orderByFields.Add($"[{matchedColumn}]");
                }
            }
            else
            {
                // 欄位不存在，嘗試使用標準欄位名直接查詢
                var directMatch = actualColumns.FirstOrDefault(c => 
                    c.Equals(canonicalField, StringComparison.OrdinalIgnoreCase));

                if (directMatch != null)
                {
                    selectClauses.Add($"[{directMatch}]");
                    
                    
                    if (canonicalField == "ClassName" || canonicalField == "SeatNumber" || canonicalField == "StudentId")
                    {
                        orderByFields.Add($"[{directMatch}]");
                    }
                }
                // 若都不存在，則跳過該欄位（後續驗證會報錯）
            }
        }

        // 若沒有任何欄位，使用 SELECT *
        if (selectClauses.Count == 0)
        {
            return $"SELECT * FROM [{sourceName}]";
        }

        var sql = $"SELECT {string.Join(", ", selectClauses)} FROM [{sourceName}]";

        // 加入排序（若有可用欄位）
        if (orderByFields.Count > 0)
        {
            sql += $" ORDER BY {string.Join(", ", orderByFields)}";
        }

        return sql;
    }

    /// <summary>
    /// 驗證資料
    /// </summary>
    public List<string> ValidateData(DataTable data)
    {
        var errors = new List<string>();

        // 檢查必要欄位是否存在
        var requiredFields = new[] { "ClassName", "ParentName", "StudentId", "Barcode1", "Barcode2", "Barcode3" };
        foreach (var field in requiredFields)
        {
            if (!data.Columns.Contains(field))
            {
                errors.Add($"缺少必要欄位：{field}");
            }
        }

        if (errors.Count > 0)
            return errors;

        // 逐筆驗證資料
        for (int i = 0; i < data.Rows.Count; i++)
        {
            var row = data.Rows[i];
            var rowNum = i + 1;

            // 必填檢核 [REQ-VAL-001]
            ValidateRequired(row, "ClassName", rowNum, errors);
            ValidateRequired(row, "StudentId", rowNum, errors);
            ValidateRequired(row, "ParentName", rowNum, errors);
            ValidateRequired(row, "Barcode1", rowNum, errors);
            ValidateRequired(row, "Barcode2", rowNum, errors);
            ValidateRequired(row, "Barcode3", rowNum, errors);

            // 格式檢核 [REQ-VAL-002]
            if (data.Columns.Contains("StudentId"))
            {
                var studentId = row["StudentId"]?.ToString() ?? "";
                if (!string.IsNullOrEmpty(studentId) && studentId.Length != 6)
                {
                    errors.Add($"第 {rowNum} 筆：學號必須為 6 碼（目前：{studentId.Length} 碼）");
                }
            }

            // 條碼字元檢核 [REQ-BC-001]
            if (data.Columns.Contains("Barcode1"))
                ValidateBarcodeCharacters(row["Barcode1"]?.ToString(), "Barcode1", rowNum, errors);
            if (data.Columns.Contains("Barcode2"))
                ValidateBarcodeCharacters(row["Barcode2"]?.ToString(), "Barcode2", rowNum, errors);
            if (data.Columns.Contains("Barcode3"))
                ValidateBarcodeCharacters(row["Barcode3"]?.ToString(), "Barcode3", rowNum, errors);
        }

        return errors;
    }

    private static void ValidateRequired(DataRow row, string fieldName, int rowNum, List<string> errors)
    {
        if (!row.Table.Columns.Contains(fieldName))
            return;

        var value = row[fieldName]?.ToString();
        if (string.IsNullOrWhiteSpace(value))
        {
            errors.Add($"第 {rowNum} 筆：{fieldName} 不得為空");
        }
    }

    /// <summary>
    /// 驗證條碼字元（Code39 允許字元）
    /// </summary>
    private static void ValidateBarcodeCharacters(string? value, string fieldName, int rowNum, List<string> errors)
    {
        if (string.IsNullOrEmpty(value))
            return;

        // Code39 允許字元：0-9, A-Z, -, ., 空白, $, /, +, %
        const string allowedChars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ-. $/+%";

        foreach (var c in value.ToUpperInvariant())
        {
            if (!allowedChars.Contains(c))
            {
                errors.Add($"第 {rowNum} 筆：{fieldName} 包含不允許的字元 '{c}'");
                break;
            }
        }
    }

    public void Dispose()
    {
        if (_disposed)
            return;

        _connection?.Close();
        _connection?.Dispose();
        _disposed = true;

        GC.SuppressFinalize(this);
    }
}
