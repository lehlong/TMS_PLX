namespace DMS.BUSINESS.Models
{
    public class CalculateDiscountOutputModel
    {
        public Dlg Dlg { get; } = new Dlg();
        public List<PtModel> Pt { get; } = new List<PtModel>();
    }

    public class Dlg
    {
        public IList<DlgModel> Dlg1 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg2 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg3 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg4 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg5 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg6 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg7 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg8 { get; } = new List<DlgModel>();
    }

    public class DlgModel
    {
        public string GoodCode { get; set; }
        public string GoodName { get; set; }
        public string LocalCode { get; set; }
        public string Stt { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
        public decimal Col11 { get; set; } = 0;
        public decimal Col12 { get; set; } = 0;
        public decimal Col13 { get; set; } = 0;
        public decimal Col14 { get; set; } = 0;
        public decimal Col15 { get; set; } = 0;
        public decimal Col16 { get; set; } = 0;
        public bool IsBold { get; set; } = false;
        public string Note { get; set; }
    }

    public class PtModel
    {
        
        public string MarketCode { get; set; }
        public string MarketName { get; set; }
        public string LocalCode { get; set; }
        public string Stt { get; set; }
        public decimal Col1 { get; set; } = 0;
        public decimal Col2 { get; set; } = 0;
        public decimal Col3 { get; set; } = 0;
        public decimal Col4 { get; set; } = 0;
        public decimal Col5 { get; set; } = 0;
        public decimal Col6 { get; set; } = 0;
        public decimal Col7 { get; set; } = 0;
        public decimal Col8 { get; set; } = 0;
        public decimal Col9 { get; set; } = 0;
        public decimal Col10 { get; set; } = 0;
        public decimal Col11 { get; set; } = 0;
        public decimal Col12 { get; set; } = 0;
        public decimal Col13 { get; set; } = 0;
        public decimal Col14 { get; set; } = 0;
        public decimal Col15 { get; set; } = 0;
        public decimal Col16 { get; set; } = 0;
        public decimal Col17 { get; set; } = 0;
        public decimal Col18 { get; set; } = 0;
        public decimal Col19 { get; set; } = 0;
        public decimal Col20 { get; set; } = 0;
        public decimal Col21 { get; set; } = 0;
        public decimal Col22 { get; set; } = 0;
        public decimal Col23 { get; set; } = 0;
        public decimal Col24 { get; set; } = 0;
        public decimal Col25 { get; set; } = 0;
        public decimal Col26 { get; set; } = 0;
        public decimal Col27 { get; set; } = 0;
        public decimal Col28 { get; set; } = 0;
        public decimal Col29 { get; set; } = 0;
        public decimal Col30 { get; set; } = 0;
        public bool IsBold { get; set; } = false;
        public string Note { get; set; }
    }
}
