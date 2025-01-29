using CoreTaxXmlWriter.Models;
using OfficeOpenXml;
using OfficeOpenXml.Export.ToCollection;

namespace CoreTaxXlsxReader;

public class TaxInvoiceReader
{
    public TaxInvoiceBulk ReadFile()
    {
        TaxInvoiceBulk result = new();
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        using (ExcelPackage package = new(@"CoreTax.xlsx"))
        {
            ExcelWorksheet invoiceSheet = package.Workbook.Worksheets["Faktur"];
            result.TIN = invoiceSheet.Cells[1, 3].Text;
            List<TaxInvoiceWrapper> invoiceList = ReadInvoiceSheet(invoiceSheet);

            ExcelWorksheet detailSheet = package.Workbook.Worksheets["DetailFaktur"];
            List<GoodServiceWrapper> detailList = ReadDetailSheet(detailSheet);

            ComposeTaxInvoiceBulk(result, invoiceList, detailList);
        }

        return result;
    }

    private List<TaxInvoiceWrapper> ReadInvoiceSheet(ExcelWorksheet worksheet)
    {
        return worksheet.Cells[4, 1, GetLastRow(worksheet), worksheet.Dimension.Columns]
            .ToCollectionWithMappings<TaxInvoiceWrapper>(
                row =>
                {
                    TaxInvoiceWrapper invoiceWrapper = new()
                    {
                        InvoiceSequenceNumber = row.GetValue<int>(0),
                    };

                    invoiceWrapper.TaxInvoice.Date = row.GetValue<DateTime>(1);
                    invoiceWrapper.TaxInvoice.TaxInvoiceOpt = row.GetText(2);
                    invoiceWrapper.TaxInvoice.TrxCode = row.GetText(3);
                    invoiceWrapper.TaxInvoice.AddInfo = row.GetText(4);
                    invoiceWrapper.TaxInvoice.CustomDoc = row.GetText(5);
                    invoiceWrapper.TaxInvoice.RefDesc = row.GetText(6);
                    invoiceWrapper.TaxInvoice.FacilityStamp = row.GetText(7);
                    invoiceWrapper.TaxInvoice.SellerIDTKU = row.GetText(8);
                    invoiceWrapper.TaxInvoice.BuyerTin = row.GetText(9);
                    invoiceWrapper.TaxInvoice.BuyerDocument = row.GetText(10);
                    invoiceWrapper.TaxInvoice.BuyerCountry = row.GetText(11);
                    invoiceWrapper.TaxInvoice.BuyerDocumentNumber = row.GetText(12);
                    invoiceWrapper.TaxInvoice.BuyerName = row.GetText(13);
                    invoiceWrapper.TaxInvoice.BuyerAdress = row.GetText(14);
                    invoiceWrapper.TaxInvoice.BuyerEmail = row.GetText(15);
                    invoiceWrapper.TaxInvoice.BuyerIDTKU = row.GetText(16);

                    return invoiceWrapper;
                },
                options =>
                {
                    options.HeaderRow = null;
                    options.DataStartRow = 0;
                }
            );
    }

    private List<GoodServiceWrapper> ReadDetailSheet(ExcelWorksheet worksheet)
    {
        return worksheet.Cells[2, 1, GetLastRow(worksheet), worksheet.Dimension.Columns]
            .ToCollectionWithMappings<GoodServiceWrapper>(
                row =>
                {
                    GoodServiceWrapper goodServiceWrapper = new()
                    {
                        InvoiceSequenceNumber = row.GetValue<int>(0),
                    };

                    goodServiceWrapper.GoodService.Opt = row.GetText(1);
                    goodServiceWrapper.GoodService.Code = row.GetText(2);
                    goodServiceWrapper.GoodService.Name = row.GetText(3);
                    goodServiceWrapper.GoodService.Unit = row.GetText(4);
                    goodServiceWrapper.GoodService.Price = decimal.Round(row.GetValue<decimal>(5), 2);
                    goodServiceWrapper.GoodService.Qty = decimal.Round(row.GetValue<decimal>(6), 2);
                    goodServiceWrapper.GoodService.TotalDiscount = decimal.Round(row.GetValue<decimal>(7), 2);
                    goodServiceWrapper.GoodService.TaxBase = decimal.Round(row.GetValue<decimal>(8), 2);
                    goodServiceWrapper.GoodService.OtherTaxBase = decimal.Round(row.GetValue<decimal>(9), 2);
                    goodServiceWrapper.GoodService.VATRate = row.GetValue<byte>(10);
                    goodServiceWrapper.GoodService.VAT = decimal.Round(row.GetValue<decimal>(11), 2);
                    goodServiceWrapper.GoodService.STLGRate = row.GetValue<byte>(12);
                    goodServiceWrapper.GoodService.STLG = decimal.Round(row.GetValue<decimal>(13), 2);

                    return goodServiceWrapper;
                },
                options =>
                {
                    options.HeaderRow = null;
                    options.DataStartRow = 0;
                }
            );
    }

    private int GetLastRow(ExcelWorksheet worksheet)
    {
        int lastRow = worksheet.Dimension.Rows;

        if (worksheet.Cells[lastRow, 1].Text.Trim() == "END")
        {
            lastRow--;
        }

        return lastRow;
    }

    private void ComposeTaxInvoiceBulk(
        TaxInvoiceBulk result,
        List<TaxInvoiceWrapper> invoices,
        List<GoodServiceWrapper> invoiceDetails)
    {
        foreach (TaxInvoiceWrapper invoice in invoices)
        {
            result.ListOfTaxInvoice.Add(invoice.TaxInvoice);
            IEnumerable<GoodServiceWrapper> filteredDetails = invoiceDetails
                .Where(d => d.InvoiceSequenceNumber == invoice.InvoiceSequenceNumber);

            foreach (GoodServiceWrapper item in filteredDetails)
            {
                invoice.TaxInvoice.ListOfGoodService.Add(item.GoodService);
            }
        }
    }
}
