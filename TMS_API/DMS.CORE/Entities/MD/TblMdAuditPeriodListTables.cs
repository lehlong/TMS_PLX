using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_AUDIT_PERIOD_LIST_TABLES")]
    public class TblMdAuditPeriodListTables : BaseEntity
    {
        [Key]
        [Column("CODE")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int Code { get; set; }
        [Column("LIST_TABLES_CODE", TypeName = "VARCHAR(50)")]
        public string ListTablesCode { get; set; }
        [Column("AUDIT_PERIOD_CODE")]
        public string AuditPeriodCode { get; set; }
        [Column("VERSION")]
        public int Version { get; set; }
        [Column("STATUS")]
        public string Status { get; set; }
        [ForeignKey("ListTablesCode")]
        public virtual TblMdTemplateListTables? TblMdTemplateListTables { get; set; }
        public List<TblMdAuditTemplateListTablesData> AuditTemplateListTablesReferences { get; set; }

    }
}
