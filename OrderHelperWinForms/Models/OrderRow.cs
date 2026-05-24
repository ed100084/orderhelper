namespace OrderHelperWinForms.Models;

public record OrderRow(
    string OrderNo,
    string ItemNo,
    string Code,
    string Name,
    string Unit,
    string Quantity,
    string Vendor,
    string Tel,
    string Fax
);
