using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.IN
{
    [Table("T_IN_HE_SO_MAT_HANG")]
    public class TblInHeSoMatHang : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { get; set; }

        [Column("GOODS_CODE", TypeName = "NVARCHAR(255)")]
        public string GoodsCode { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { set; get; }

        [Column("HE_SO_VCF", TypeName = "DECIMAL(18,4)")]
        public decimal HeSoVcf { get; set; }

        [Column("THUE_BVMT", TypeName = "DECIMAL(18,0)")]
        public decimal ThueBvmt { get; set; }

        [Column("L15_CHUA_VAT_BVMT", TypeName = "DECIMAL(18,0)")]
        public decimal L15ChuaVatBvmt { get; set; }

        [Column("L15_CHUA_VAT_BVMT_NBL", TypeName = "DECIMAL(18,0)")]
        public decimal L15ChuaVatBvmtNbl { get; set; }

        [Column("GIAM_GIA_FOB", TypeName = "DECIMAL(18,0)")]
        public decimal GiamGiaFob { get; set; }

        [Column("LAI_GOP_DIEU_TIET", TypeName = "DECIMAL(18,0)")]
        public decimal LaiGopDieuTiet { get; set; }

        [Column("FROM_DATE", TypeName = "DATETIME")]
        public DateTime? FromDate { get; set; }

        [Column("TO_DATE", TypeName = "DATETIME")]
        public DateTime? ToDate { get; set; }
    }
}
