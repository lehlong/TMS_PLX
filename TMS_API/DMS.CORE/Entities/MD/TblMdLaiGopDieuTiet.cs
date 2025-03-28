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
    [Table("T_MD_LAI_GOP_DIEU_TIET")]
    public class TblMdLaiGopDieuTiet : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(1000)")]
        public string Code { set; get; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]    
        public string GoodsCode { set; get; }

        [Column("MARKET_CODE", TypeName = "VARCHAR(50)")]
        public string MarketCode { set; get; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string HeaderCode { set; get; } = "";

        [Column("FROM_DATE", TypeName = "DATETIME")]
        public DateTime? FromDate { set; get; }

        [Column("TO_DATE", TypeName = "DATETIME")]
        public DateTime ToDate { set; get; }

        [Column("PRICE", TypeName = "DECIMAL(18,0)")]
        public decimal? Price { set; get; }

    }
}
