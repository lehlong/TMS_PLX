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
    [Table("T_MD_MG_OPINION_LIST")]
    public class TblMdMgOpinionList : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }
        [Column("NAME", TypeName = "NVARCHAR(MAX)")]
        public string Name { get; set; }
        [Column("DESCRIPTION", TypeName = "NVARCHAR(MAX)")]
        public string Description { get; set; }
        [Column("TIME_YEAR")]
        public string TimeYear { get; set; }
        [Column("AUDIT_PERIOD")]
        public string AuditPeriod { get; set; }
        [ForeignKey("AuditPeriod")]
        public virtual TblMdAuditPeriod AuditPeriods { get; set; }
        public virtual List<tblBuOpinionList> OpinionLists { get; set; }
    }
}
