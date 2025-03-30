using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_INPUT_CUSTOMER_DB")]
    public class TblBuInputCustomerDb : SoftDeleteEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }
        [Column("HEADER_ID", TypeName = "NVARCHAR(50)")]
        public string HeaderId { get; set; }
        [Column("CODE", TypeName = "NVARCHAR(50)")]
        public string Code { get; set; }

        [Column("NAME", TypeName = "NVARCHAR(1000)")]
        public string Name { get; set; }
        [Column("LOCAL_CODE", TypeName = "NVARCHAR(50)")]
        public string LocalCode { get; set; }
        [Column("LOCAL2", TypeName = "NVARCHAR(500)")]
        public string? Local2 { get; set; }
        [Column("MARKET_CODE", TypeName = "NVARCHAR(50)")]
        public string MarketCode { get; set; }
        [Column("CU_LY_BQ", TypeName = "DECIMAL(18,0)")]
        public decimal CuLyBq { get; set; }
        [Column("CPCCVC", TypeName = "DECIMAL(18,0)")]
        public decimal Cpccvc { get; set; }

        [Column("CVCBQ", TypeName = "DECIMAL(18,0)")]
        public decimal Cvcbq { get; set; }

        [Column("LVNH", TypeName = "DECIMAL(18,0)")]
        public decimal Lvnh { get; set; }

        [Column("HTCVC", TypeName = "DECIMAL(18,0)")]
        public decimal Htcvc { get; set; }
        [Column("HTT_VB1370_PLX", TypeName = "DECIMAL(18,0)")]
        public decimal HttVb1370 { get; set; }

        [Column("CKV2", TypeName = "DECIMAL(18,0)")]
        public decimal Ckv2 { get; set; }
        [Column("PHUONG_THUC", TypeName = "NVARCHAR(50)")]
        public string PhuongThuc { get; set; }

        [Column("THTT", TypeName = "NVARCHAR(50)")]
        public string Thtt { get; set; }
        [Column("ADDRESS", TypeName = "NVARCHAR(1000)")]
        public string Adrress { get; set; }
        [Column("CK_XANG", TypeName = "DECIMAL(18,0)")]
        public decimal CkXang { get; set; }
        [Column("CK_DAU", TypeName = "DECIMAL(18,0)")]
        public decimal CkDau { get; set; }

        [Column("C_ORDER")]
        public int Order { get; set; }
    }
}
