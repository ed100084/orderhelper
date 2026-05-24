namespace OrderHelperWinForms.Models;

public class HospitalSettings
{
    // Header
    public string HospitalName     { get; set; } = "義大醫療財團法人義大醫院";
    public string FormTitle        { get; set; } = "藥品訂購單";

    // Invoice block (left side)
    public string InvoiceHeader    { get; set; } = "義大醫療財團法人義大醫院";
    public string InvoiceAddress   { get; set; } = "高雄市燕巢區角宿里義大路1號";
    public string TaxId            { get; set; } = "25886456";
    public string MedicalCode      { get; set; } = "1142120001";
    public string DrugLicenseNo    { get; set; } = "QHP101000003";

    // Delivery block (right side of invoice section)
    public string DeliveryAddress  { get; set; } = "高雄市燕巢區角宿里義大路1號";
    public string DeliveryNote     { get; set; } = "□藥庫(B1F)   □_______";
    public string ContactPhone     { get; set; } = "07-6150011#6226.6225.6224(藥庫)林藥師";
    public string ContactFax       { get; set; } = "07-6154431";

    // Notes (right column)
    public string Note1 { get; set; } = "1.發票與出貨單請隨貨附上。";
    public string Note2 { get; set; } = "2.請開立三聯式發票。";
    public string Note3 { get; set; } = "3.請附8元回郵信封(郵寄折讓單)。";
    public string Note4 { get; set; } = "4.首次交易或未申請匯款者請與採購課人員接洽辦理。";

    public static HospitalSettings Default() => new();
}
