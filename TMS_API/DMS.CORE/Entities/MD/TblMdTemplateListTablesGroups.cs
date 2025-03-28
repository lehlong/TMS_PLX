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
    [Table("T_MD_TEMPLATE_LIST_TABLES_GROUPS")]
    public class TblMdTemplateListTablesGroups : SoftDeleteEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("CODE")]
        public Guid Code { get; set; }
        [Column("ID", TypeName = "VARCHAR(50)")]
        public string? Id { get; set; }
        [Column("NAME", TypeName = "NVARCHAR(MAX)")]
        public string? Name { get; set; }
        public virtual List<TblMdTemplateListTables> TemplateListTables { get; set; }
    }
}
