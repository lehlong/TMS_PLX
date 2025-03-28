using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DMS.CORE.Entities.AD;

namespace DMS.CORE.Entities.BU
{
    [Table("T_CM_COMMENT")]
    public class TblCmComment : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public int Id { get; set; }

        [Column("P_ID")]
        public int? PId { get; set; }

        [Column("TYPE",TypeName = "varchar2(255)")]
        public string Type { get; set; }

        [Column("CONTENT",TypeName = "nvarchar2(1000)")]
        public string Content { get; set; }

        [Column("ATTACHMENT_ID")]
        public Guid? AttachmentId { get; set; }

        [ForeignKey("PId")]
        public virtual TblCmComment PComment { get; set; }

        public virtual List<TblCmComment> Replies { get; set; }

        [ForeignKey("AttachmentId")]
        public virtual TblCmAttachment Attachment { get; set; }

        public virtual TblCmModuleComment ModuleComment { get; set; }

        [ForeignKey("CreateBy")]
        public TblAdAccount Creator { get; set; }

        [ForeignKey("UpdateBy")]
        public TblAdAccount Updater { get; set; }
    }
}
