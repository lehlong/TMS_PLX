using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.BU;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_TEMPLATE_LIST_TABLES")]
    public class TblMdTemplateListTables : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("NAME", TypeName = "NVARCHAR(MAX)")]
        public string? Name { get; set; }

        [Column("TIME_YEAR")]
        public string? TimeYear { get; set; }

        [Column("NOTE", TypeName = "NVARCHAR(350)")]
        public string? Note { get; set; }
        [Column("ORG_CODE", TypeName = "VARCHAR(50)")]
        public string? OrgCode { get; set; }
        [Column("MG_LIST_TABLES_CODE", TypeName = "VARCHAR(50)")]
        public string? MgListTablesCode { get; set; }
        [Column("GROUP_CODE")]
        public Guid? GroupCode { get; set; }
        [ForeignKey("GroupCode")]
        public virtual TblMdTemplateListTablesGroups? TemplateListTablesGroups { get; set; }
        public virtual List<TblMdTemplateListTablesData> TemDataReferences { get; set; }
    }

}
