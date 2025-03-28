using DMS.CORE.Common;

using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_TEMPLATE_REPORT")]
    public class TblMdTemplateReport : BaseEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }

        [Column("YEAR", TypeName = "NVARCHAR(50)")]
        public string? Year { get; set; }

        [Column("AUDIT_PERIOD", TypeName = "NVARCHAR(50)")]
        public string? AuditPeriod { get; set; }

        [Column("FILE_NAME", TypeName = "NVARCHAR(2000)")]
        public string? FileName { get; set; }

        [Column("FILE_OLD_NAME", TypeName = "NVARCHAR(2000)")]
        public string? FileOldName { get; set; }

        [Column("FILE_EXT", TypeName = "NVARCHAR(1000)")]
        public string? FileExt { get; set; }

        [Column("FILE_SIZE", TypeName = "NVARCHAR(50)")]
        public string? FileSize { get; set; }

        [Column("NETWORK_PATH", TypeName = "NVARCHAR(500)")]
        public string? NetworkPath { get; set; }

        [Column("FULL_PATH", TypeName = "NVARCHAR(2000)")]
        public string? FullPath { get; set; }

    }
}
