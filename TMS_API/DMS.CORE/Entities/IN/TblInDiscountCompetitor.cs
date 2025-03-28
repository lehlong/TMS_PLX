using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.IN
{
    [Table("T_IN_DISCOUNT_COMPETITOR")]
    public class TblInDiscountCompetitor : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { get; set; }

        [Column("COMPETITOR_CODE", TypeName = "NVARCHAR(255)")]
        public string CompetitorCode { get; set; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]
        public string GoodsCode { set; get; }

        [Column("DISCOUNT", TypeName = "DECIMAL(18,0)")]
        public decimal? Discount { get; set; }

    }
}
