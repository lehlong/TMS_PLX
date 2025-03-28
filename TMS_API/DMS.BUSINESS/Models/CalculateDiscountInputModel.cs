using DMS.CORE.Entities.BU;

namespace DMS.BUSINESS.Models
{
    public class CalculateDiscountInputModel
    {
        public TblBuCalculateDiscount Header { get; set; } = new TblBuCalculateDiscount();
        public List<TblBuInputPrice> InputPrice { get; set; } = new List<TblBuInputPrice>();
        public List<TblBuInputMarket> Market { get; set; } = new List<TblBuInputMarket>();
        public List<TblBuInputCustomerDb> CustomerDb { get; set; } = new List<TblBuInputCustomerDb>();
    }
}
