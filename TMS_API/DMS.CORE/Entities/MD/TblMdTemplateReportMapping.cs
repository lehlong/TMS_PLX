using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_TEMPLATE_REPORT_MAPPING")]
    public class TblMdTemplateReportMapping : BaseEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }

        [Column("FILE_ID", TypeName = "NVARCHAR(50)")]
        public string FileId { get; set; }

        [Column("TEXT_ELEMENT", TypeName = "NVARCHAR(50)")]
        public string? TextElement { get; set; }

        [Column("TYPE", TypeName = "NVARCHAR(50)")]
        public string? Type { get; set; }

        [Column("VALUE_INPUT", TypeName = "NVARCHAR(50)")]
        public string? ValueInput { get; set; }

        [Column("ORG_CODE", TypeName = "NVARCHAR(50)")]
        public string? OrgCode { get; set; }

        [Column("OPINION_CODE", TypeName = "NVARCHAR(50)")]
        public string? OpinionCode { get; set; }

        [Column("TABLE_CODE", TypeName = "NVARCHAR(50)")]
        public string? TableCode { get; set; }

    }
}
