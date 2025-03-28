using DMS.COMMON.Common.Enum;
using DMS.CORE.Common;
using DMS.CORE.Entities.BU;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_LIST_AUDIT")]
    public class TblMdListAudit : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }
        [Column("NAME", TypeName = "NVARCHAR(250)")]
        public string? Name { get; set; }
        [Column("TIME_YEAR", TypeName = "VARCHAR(10)")]
        public string? TimeYear { get; set; }
        [Column("AUDIT_PERIOD", TypeName = "NVARCHAR(50)")]
        public string? AuditPeriod { get; set; }
        [Column("REPORT_DATE", TypeName = "DATETIME")]
        public DateTime? ReportDate { get; set; }
        [Column("REPORT_NUMBER", TypeName ="VARCHAR(50)")]
        public string? ReportNumber { get; set; }
        [Column("STATUS", TypeName = "NVARCHAR(50)")]
        public string? Status { get; set; }
        [Column("START_DATE", TypeName = "DATETIME")]
        public DateTime? StartDate { get; set; }
        [Column("END_DATE", TypeName = "DATETIME")]
        public DateTime? EndDate { get; set; }
        [Column("OPINION_CODE", TypeName = "VARCHAR(50)")]
        public string? OpinionCode { get; set; }
        [Column("NOTE", TypeName = "NVARCHAR(250)")]
        public string? Note { get; set; }
        [Column("FILE_ID", TypeName = "UNIQUEIDENTIFIER")]
        public Guid? FileId { get; set; }
        [Column("TEXT_CONTENT", TypeName = "NVARCHAR(250)")]
        public string? TextContent { get; set; }
        [Column("APPROVER", TypeName = "NVARCHAR(50)")]
        public string? Approver { get; set; }

    }
}
