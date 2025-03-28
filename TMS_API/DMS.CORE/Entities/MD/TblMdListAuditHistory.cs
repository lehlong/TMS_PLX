using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_LIST_AUDIT_HISTORY")]
    public class tblMdListAuditHistory : SoftDeleteEntity
    {
        [Key]
        [Column("ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public string Id { get; set; }


        [Column("LIST_AUDIT_CODE", TypeName = "NVARCHAR(50)")]
        public string? ListAuditCode { get; set; }


        [Column("ACTION", TypeName = "NVARCHAR(50)")]
        public string? Action { get; set; }


        [Column("TEXT_CONTENT", TypeName = "NVARCHAR(50)")]
        public string? TextContent { get; set; }

    }
}
