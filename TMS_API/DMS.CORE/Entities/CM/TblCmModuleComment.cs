using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace DMS.CORE.Entities.BU
{
    [Table("T_CM_MODULE_COMMENT")]
    public class TblCmModuleComment : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public int Id { get; set; }

        [Column("REFERENCE_ID")]
        public Guid ReferenceId { get; set; }

        [Column("COMMENT_ID")]
        public int CommentId { get; set; }

        [Column("MODULE_TYPE",TypeName = "varchar2(50)")]
        public string? ModuleType { get; set; }

        [ForeignKey("CommentId")]
        public virtual TblCmComment Comment { get; set; }
    }
}
