using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_INPUT_PRICE")]
    public class TblBuInputPrice : SoftDeleteEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }
        [Column("HEADER_ID", TypeName = "NVARCHAR(50)")]
        public string HeaderId { get; set; }

        [Column("GOOD_CODE", TypeName = "NVARCHAR(50)")]
        public string GoodCode { get; set; }

        [Column("GOOD_NAME", TypeName = "NVARCHAR(50)")]
        public string GoodName { get; set; }

        [Column("THUE_BVMT", TypeName = "decimal(18, 0)")]
        public decimal ThueBvmt { get; set; }

        [Column("VCF", TypeName = "decimal(18, 4)")]
        public decimal Vcf { get; set; }

        [Column("CHENH_LECH", TypeName = "decimal(18, 0)")]
        public decimal ChenhLech { get; set; }

        [Column("GBL_V1", TypeName = "decimal(18, 0)")]
        public decimal GblV1 { get; set; }
        [Column("GBL_V2", TypeName = "decimal(18, 0)")]
        public decimal GblV2 { get; set; }

        [Column("L15_BLV2", TypeName = "decimal(18, 0)")]
        public decimal L15Blv2 { get; set; }

        [Column("L15_NBL", TypeName = "decimal(18, 0)")]
        public decimal L15Nbl { get; set; }
        [Column("LAI_GOP", TypeName = "decimal(18, 0)")]
        public decimal LaiGop { get; set; }

        [Column("FOB_V1", TypeName = "decimal(18, 0)")]
        public decimal FobV1 { get; set; }

        [Column("FOB_V2", TypeName = "decimal(18, 0)")]
        public decimal FobV2 { get; set; }
        [Column("C_ORDER")]
        public int Order { get; set; }

    }
}
