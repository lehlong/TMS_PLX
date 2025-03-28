using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_TEMPLATE_REPORT_ELEMENT")]
    public class TblMdTemplateReportElement : BaseEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }

        [Column("FILE_ID", TypeName = "NVARCHAR(50)")]
        public string FileId { get; set; }

        [Column("TEXT_ELEMENT", TypeName = "NVARCHAR(50)")]
        public string TextElement { get; set; }

    }
}
