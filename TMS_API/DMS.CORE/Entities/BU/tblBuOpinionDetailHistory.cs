using DMS.CORE.Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_OPINION_DETAIL_HISTORY")]
    public class tblBuOpinionDetailHistory : SoftDeleteEntity
    {
        [Key]
        [Column("ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public string Id { get; set; }


        [Column("OPINION_ID", TypeName = "NVARCHAR(50)")]
        public string? OpinionId { get; set; }


        [Column("ACTION", TypeName = "NVARCHAR(50)")]
        public string? Action { get; set; }


        [Column("TEXT_CONTENT", TypeName = "NVARCHAR(50)")]
        public string? TextContent { get; set; }

    }
}
