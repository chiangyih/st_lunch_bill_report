namespace st_lunch_bill_report.Models;

/// <summary>
/// 繳費單資料模型
/// </summary>
public class PaymentSlipData
{
    // 基本資訊
    public string 科班別 { get; set; } = string.Empty;
    public string 繳款人 { get; set; } = string.Empty;
    public string 學號 { get; set; } = string.Empty;
    public string 座號 { get; set; } = string.Empty;

    // 金額明細
    public int 學期午餐費用 { get; set; }
    public int 新生訓練餐費 { get; set; }
    public int 暑期輔導餐費 { get; set; }
    public int 退前一學期餐費 { get; set; }
    public int 超商代收費 { get; set; }
    public int 應繳金額 { get; set; }
    public int 金額 { get; set; }
    public string 實繳金額中文 { get; set; } = string.Empty;

    // 銷帳與條碼
    public string 繳費期限 { get; set; } = string.Empty;
    public string 銷帳編號 { get; set; } = string.Empty;
    public string 第1段條碼 { get; set; } = string.Empty;
    public string 第2段條碼 { get; set; } = string.Empty;
    public string 第3段條碼 { get; set; } = string.Empty;

    // 條碼圖片（由程式產生）
    public byte[]? 第1段條碼_Img { get; set; }
    public byte[]? 第2段條碼_Img { get; set; }
    public byte[]? 第3段條碼_Img { get; set; }
}
