using CoreTaxXmlWriter.Models;

public class TaxInvoiceWrapper
{
    public int InvoiceSequenceNumber { get; set; }

    public TaxInvoice TaxInvoice
    {
        get
        {
            return _taxInvoice;
        }
        set
        {
            _taxInvoice = value;
        }
    }

    private TaxInvoice _taxInvoice = new();
}