namespace OrderHelperWinForms.Models;

public record VendorPage(
    string Vendor,
    string Tel,
    string Fax,
    int PageNo,
    int TotalPages,
    List<OrderRow> Orders
);
