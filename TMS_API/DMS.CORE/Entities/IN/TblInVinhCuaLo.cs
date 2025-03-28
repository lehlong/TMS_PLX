using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.IN
{
    [Table("T_IN_VINH_CUA_LO")]
    public class TblInVinhCuaLo : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("GOODS_CODE", TypeName = "NVARCHAR(255)")]
        public string GoodsCode { get; set; }

        [Column("GBLCS_V1", TypeName = "DECIMAL(18,0)")]
        public decimal GblcsV1 { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string HeaderCode { set; get; }

        [Column("GBL_V2", TypeName = "DECIMAL(18,0)")]
        public decimal GblV2 { get; set; }

        [Column("V2_V1", TypeName = "DECIMAL(18,0)")]
        public decimal V2_V1 { get; set; }

        [Column("MTSV1", TypeName = "DECIMAL(18,0)")]
        public decimal MtsV1 { get; set; }

        [Column("GNY", TypeName = "DECIMAL(18,0)")]
        public decimal Gny { get; set; }

        [Column("CLGBLV", TypeName = "DECIMAL(18,0)")]
        public decimal Clgblv { get; set; }

        [Column("FROM_DATE", TypeName = "DATETIME")]
        public DateTime? FromDate { get; set; }

        [Column("TO_DATE", TypeName = "DATETIME")]
        public DateTime? ToDate { get; set; }
    }
}
