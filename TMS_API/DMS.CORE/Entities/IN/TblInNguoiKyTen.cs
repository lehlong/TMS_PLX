using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.IN
{
    [Table("T_IN_NGUOI_KY_TEN")]
    public class TblInNguoiKyTen : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { set; get; }

        [Column("DAI_DIEN", TypeName = "NVARCHAR(250)")]
        public string? DaiDien { get; set; }

        [Column("NGUOI_DAI_DIEN", TypeName = "NVARCHAR(250)")]
        public string? NguoiDaiDien { get; set; }

        [Column("QUYET_DINH_SO", TypeName = "NVARCHAR(250)")]
        public string? QuyetDinhSo { get; set; }
    }
}
