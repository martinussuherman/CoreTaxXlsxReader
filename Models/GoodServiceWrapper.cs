using CoreTaxXmlWriter.Models;

public class GoodServiceWrapper
{
    public int InvoiceSequenceNumber { get; set; }

    public GoodService GoodService
    {
        get
        {
            return _goodService;
        }
        set
        {
            _goodService = value;
        }
    }

    private GoodService _goodService = new();
}