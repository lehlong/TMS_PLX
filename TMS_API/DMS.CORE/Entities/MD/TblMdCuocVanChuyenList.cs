using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_CUOC_VAN_CHUYEN_LIST")]
    public class TblMdCuocVanChuyenList : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { set; get; }

        [Column("NAME", TypeName = "NVARCHAR(250)")]
        public string? Name { get; set; }
    }
}
