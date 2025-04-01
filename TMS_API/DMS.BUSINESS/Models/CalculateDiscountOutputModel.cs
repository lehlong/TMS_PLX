namespace DMS.BUSINESS.Models
{
    public class CalculateDiscountOutputModel
    {
        public Dlg Dlg { get; } = new Dlg();
        public List<DataModel> Pt { get; } = new List<DataModel>();
        public List<DataModel> Db { get; } = new List<DataModel>();
        public List<DataModel> Fob { get; } = new List<DataModel>();
        public List<DataModel> Pt09 { get; } = new List<DataModel>();
        public List<DataModel> Bbdo { get; } = new List<DataModel>();
        public List<DataModel> Pl1 { get; } = new List<DataModel>();
        public List<DataModel> Pl2 { get; } = new List<DataModel>();
        public List<DataModel> Pl3 { get; } = new List<DataModel>();
        public List<DataModel> Pl4 { get; } = new List<DataModel>();
        public List<VK11Model> Vk11Pt { get; } = new List<VK11Model>();
        public List<VK11Model> Vk11Db { get; } = new List<VK11Model>();
        public List<VK11Model> Vk11Fob { get; } = new List<VK11Model>();
        public List<VK11Model> Vk11Tnpp { get; } = new List<VK11Model>();
        public List<VK11Model> Vk11Bb { get; } = new List<VK11Model>();
        public List<VK11Model> Summary { get; } = new List<VK11Model>();
    }

    public class Dlg
    {
        public IList<DlgModel> Dlg1 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg2 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg3 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg4 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg5 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg6 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg6Old { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg7 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg8 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg9 { get; } = new List<DlgModel>();
        public IList<DlgModel> Dlg10 { get; } = new List<DlgModel>();
        public IList<DlgModel> DlgTDGBL { get; } = new List<DlgModel>();
        public IList<DlgModel> DlgTdGgptbl { get; } = new List<DlgModel>();
        public string? NameOld { get; set; }
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

    public class DataModel
    {
        public string Id { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string MarketCode { get; set; }
        public string MarketName { get; set; }
        public string LocalCode { get; set; }
        public string DeliveryPoint { get; set; }
        public string GoodCode { get; set; }
        public string GoodName { get; set; }
        public string PThuc { get; set; }
        public string Dvt { get; set; }
        public string TToan { get; set; }
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
        public decimal Col31 { get; set; } = 0;
        public decimal Col32 { get; set; } = 0;
        public decimal Col33 { get; set; } = 0;
        public decimal Col34 { get; set; } = 0;
        public bool IsBold { get; set; } = false;
        public string Note { get; set; }
    }
    public class VK11Model
    {
        public string Stt { get; set; }
        public bool IsBold { get; set; } = false;
        public string CustomerName { get; set; }
        public string GoodsName { get; set; }
        public string Address { get; set; }
        public string MarketCode { get; set; }
        public string MarketName { get; set; }
        public decimal Col1 { get; set; }
        public decimal Col2 { get; set; }
        public string Col3 { get; set; }
        public string Col4 { get; set; }
        public string Col5 { get; set; }
        public string Col6 { get; set; } = "L";
        public string Col7 { get; set; }
        public decimal Col8 { get; set; }
        public string Col9 { get; set; } = "VND";
        public int Col10 { get; set; } = 1;
        public string Col11 { get; set; } = "L";
        public string Col12 { get; set; } = "C";
        public string Col13 { get; set; }
        public string Col14 { get; set; }
        public string Col15 { get; set; } = "31.12.9999";
    }
}
