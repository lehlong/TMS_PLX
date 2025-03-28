using DMS.CORE.Common;
using DMS.CORE.Entities.AD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_AUDIT_PERIOD")]
    public class TblMdAuditPeriod : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "NVARCHAR(50)")]
        public string Code { get; set; }
        [Column("AUDIT_PERIOD",TypeName = "NVARCHAR(MAX)")]
        public string AuditPeriod { get; set; }
        public virtual List<TblMdMgListTables> MdMgListTables { get; set; }
    }
}
