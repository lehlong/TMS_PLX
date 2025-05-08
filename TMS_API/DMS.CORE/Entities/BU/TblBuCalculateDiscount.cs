using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_CALCULATE_DISCOUNT")]
    public class TblBuCalculateDiscount : SoftDeleteEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }

        [Column("NAME", TypeName = "NVARCHAR(500)")]
        public string Name { get; set; }

        [Column("DATE", TypeName = "DATETIME")]
        public DateTime Date { get; set; }

        [Column("HOUR", TypeName = "DATETIME")]
        public DateTime? Hour { get; set; }

        [Column("SIGNER_CODE", TypeName = "NVARCHAR(50)")]
        public string? SignerCode { get; set; }

        [Column("TCKT_CODE", TypeName = "NVARCHAR(50)")]
        public string? TcktCode { get; set; }

        [Column("KDXD_CODE", TypeName = "NVARCHAR(50)")]
        public string? KdxdCode { get; set; }

        [Column("VIETPHUONGAN_CODE", TypeName = "NVARCHAR(50)")]
        public string? VietphuonganCode { get; set; }

        [Column("QUYET_DINH_SO", TypeName = "NVARCHAR(250)")]
        public string? QuyetDinhSo { get; set; }

        [Column("CONG_DIEN_SO", TypeName = "NVARCHAR(250)")]
        public string? CongDienSo { get; set; }
        [Column("CONG_DIEN_PT_BAN_LE", TypeName = "NVARCHAR(250)")]
        public string? CongDienPtBanLe { get; set; }

        [Column("STATUS", TypeName = "NVARCHAR(50)")]
        public string? Status { get; set; }
    }
}
