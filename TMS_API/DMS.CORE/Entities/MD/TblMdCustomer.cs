using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_CUSTOMER")]
    public class TblMdCustomer : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { set; get; }

        [Column("NAME", TypeName = "NVARCHAR(255)")]
        public string Name { set; get; }

        //[Column("PHONE", TypeName = "VARCHAR(50)")]
        //public string? Phone { set; get; }

        //[Column("EMAIL", TypeName = "NVARCHAR(255)")]
        //public string? Email { set; get; }

        [Column("ADDRESS", TypeName = "NVARCHAR(255)")]
        public string? Address { set; get; }

        [Column("PAYMENT_TERM", TypeName = "NVARCHAR(255)")]
        public string? PaymentTerm { set; get; }

        [Column("GAP", TypeName = "DECIMAL(18,0)")]
        public decimal? Gap { set; get; }

        [Column("HO_TRO_CUOC_VC", TypeName = "DECIMAL(18,0)")]
        public decimal? HoTroCuocVc { set; get; }

        [Column("CUOC_VC_BQ", TypeName = "DECIMAL(18,0)")]
        public decimal? CuocVcBq { set; get; }

        [Column("BUY_INFO", TypeName = "NVARCHAR(500)")]
        public string? BuyInfo { set; get; }

        [Column("BANK_LOAN_INTEREST", TypeName = "DECIMAL(18,0)")]
        public decimal? BankLoanInterest { set; get; }

        [Column("SALES_METHOD_CODE", TypeName = "VARCHAR(50)")]
        public string? SalesMethodCode { set; get; }

        [Column("LOCAL_CODE", TypeName = "VARCHAR(50)")]
        public string? LocalCode { set; get; }

        [Column("MARKET_CODE", TypeName = "VARCHAR(50)")]
        public string? MarketCode { set; get; }

        [Column("CUSTOMER_TYPE_CODE", TypeName = "VARCHAR(50)")]
        public string? CustomerTypeCode { set; get; }

        [Column("MGGLH_XANG", TypeName = "DECIMAL(18,0)")]
        public decimal? MgglhXang { set; get; }

        [Column("MGGLH_DAU", TypeName = "DECIMAL(18,0)")]
        public decimal? MgglhDau { set; get; }

        [Column("FOB", TypeName = "DECIMAL(18,0)")]
        public decimal? Fob { set; get; }

    }
}
