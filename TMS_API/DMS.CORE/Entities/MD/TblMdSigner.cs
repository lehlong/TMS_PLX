using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DMS.CORE.Common;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_SIGNER")]
    public class TblMdSigner : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }
        [Column("NAME", TypeName = "NVARCHAR(255)")]
        public string Name { get; set; }

        [Column("POSITION", TypeName = "NVARCHAR(255)")]
        public string Position { get; set; }
    }
}
