using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.IN
{
    [Table("T_IN_DISCOUNT_COMPANY")]
    public class TblInDiscountCompany
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { set; get; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]
        public string? GoodsCode { set; get; }

        [Column("DISCOUNT", TypeName = "DECIMAL(18, 0)")]
        public decimal? Discount { get; set; }

    }
}
