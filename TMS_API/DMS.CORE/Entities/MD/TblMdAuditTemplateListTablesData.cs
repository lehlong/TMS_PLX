using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_AUDIT_TEMPLATE_LIST_TABLES_DATA")]
    public class TblMdAuditTemplateListTablesData
    {
        [Column("TEMPLATE_DATA_CODE")]
        public Guid TemplateDataCode { get; set; }

        [ForeignKey("TemplateDataCode")]
        public virtual TblMdTemplateListTablesData templateListTablesData { get; set; }

        [Column("AUDIT_LIST_TABLES_CODE")]
        public int AuditListTablesCode { get; set; }

        [ForeignKey("AuditListTablesCode")]
        public virtual TblMdAuditPeriodListTables auditPeriodListTables { get; set; }
        [Column("UNIT")]
        public string? Unit { get; set; }

        [Column("AUDIT_VALUE")]
        public decimal? AuditValue { get; set; }

        [Column("AUDIT_EXPLANATION")]
        public string? AuditExplanation { get; set; }

        [Column("EXPLANATION_VALUE")]
        public decimal? ExplanationValue { get; set; }

        [Column("EXPLANATION_NOTE")]
        public string? ExplanationNote { get; set; }

    }
}
