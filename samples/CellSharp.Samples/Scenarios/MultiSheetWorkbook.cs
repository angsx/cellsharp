using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class MultiSheetWorkbook
{
    internal static void Run(string outputDirectory)
    {
        var customerSchema = Excel.Schema<Customer>()
            .SheetName("Customers")
            .Column(x => x.Id, column => column.Header("Customer ID").Width(14))
            .Column(x => x.Name, column => column.Header("Company").Width(28))
            .Column(x => x.CreatedAt, column => column.Header("Created").Format("yyyy-mm-dd").Width(14))
            .Build();
        var orderSchema = Excel.Schema<OrderLine>()
            .SheetName("Orders")
            .AsTable("OrdersTable")
            .FreezePanes(1, 0)
            .Column(x => x.Product, column => column.Width(24))
            .Column(x => x.Quantity, column => column.Format("0").Width(10))
            .Column(x => x.UnitPrice, column => column.Header("Unit price").Format("#,##0.00").Width(14))
            .Column(x => x.Total, column => column.Header("Total").Format("#,##0.00").Width(14))
            .Build();
        var path = Path.Combine(outputDirectory, "customer-orders.xlsx");

        Excel.Workbook()
            .AddSheet(new[] { new Customer { Id = 1, Name = "Acme Ltd", CreatedAt = new DateTime(2025, 1, 15) } }, customerSchema, options => options.FreezeHeaderRow())
            .AddSheet(new[] { new OrderLine { Product = "USB-C Dock", Quantity = 2, UnitPrice = 129.00m, Total = 258.00m } }, orderSchema, options => options.AlternatingRows())
            .Write(path);

        using var workbook = Excel.Open(path);
        var customers = workbook.Read(customerSchema);
        var orders = workbook.Read(orderSchema);
        Console.WriteLine($"  Read {customers.Rows.Count} customer(s) and {orders.Rows.Count} order(s) from one workbook.");
    }
}
