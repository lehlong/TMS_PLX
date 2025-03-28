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
    [Table("T_MD_GIA_GIAO_TAP_DOAN")]
    public class TblMdGiaGiaoTapDoan : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { set; get; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]
        public string? GoodsCode { set; get; }

        [Column("GGTDL_CODE", TypeName = "VARCHAR(50)")]
        public string? GgtdlCode { set; get; }

        [Column("OLD_PRICE", TypeName = "DECIMAL(18.0)")]
        public decimal? OldPrice { set; get; }

        [Column("NEW_PRICE", TypeName = "DECIMAL(18.0)")]
        public decimal? NewPrice { set; get; }

    }
}
