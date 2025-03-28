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
    [Table("T_MD_AUDIT_TEMPLATE_HISTORY")]
    public class TblMdAuditTemplateHistory : SoftDeleteEntity
    {
        [Key]
        [Column("ID")]
        public string Id { get; set; }


        [Column("LIST_AUDIT_CODE", TypeName = "NVARCHAR(50)")]
        public string? ListAuditCode { get; set; }


        [Column("ACTION", TypeName = "NVARCHAR(50)")]
        public string? Action { get; set; }


        [Column("TEXT_CONTENT", TypeName = "NVARCHAR(50)")]
        public string? TextContent { get; set; }  

        [Column("AUDIT_PERIOD_CODE")]
        public int? AuditPeriodListTablesCode { get; set; }



    }
}
