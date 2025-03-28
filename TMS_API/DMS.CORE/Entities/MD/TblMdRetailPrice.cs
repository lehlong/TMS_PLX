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
    [Table("T_MD_RETAIL_PRICE")]
    public class TblMdRetailPrice : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { set; get; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]    
        public string GoodsCode { set; get; }

        [Column("GBLL_CODE", TypeName = "VARCHAR(50)")]
        public string? GbllCode { set; get; }

        [Column("OLD_PRICE", TypeName = "DECIMAL(18,0)")]
        public decimal OldPrice { set; get; }

        [Column("NEW_PRICE", TypeName = "DECIMAL(18,0)")]
        public decimal NewPrice { set; get; }

        [Column("MUC_TANG_V1", TypeName = "DECIMAIL(18,0)")]
        public decimal MucTangV1 { set; get; }

    }
}
