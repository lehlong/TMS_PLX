using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace DMS.CORE.Entities.BU
{
    [Table("T_CM_MODULE_ATTACHMENT")]
    public class TblCmModuleAttachment : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public Guid Id { get; set; }

        [Column("REFERENCE_ID")]
        public Guid? ReferenceId { get; set; }

        [Column("MODULE_TYPE", TypeName = "VARCHAR(50)")]
        public string ModuleType { get; set; }

        [Column("ATTACHMENT_ID")]
        public Guid AttachmentId { get; set; }

        [ForeignKey("AttachmentId")]
        public virtual TblCmAttachment Attachment { get; set; }
    }
}
