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
    [Table("T_BU_OPINION_DETAIL")]
    public class tblBuOpinionDetail : SoftDeleteEntity
    {
        [Key]
        [Column("ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public string Id { get; set; }


        [Column("MG_CODE", TypeName = "NVARCHAR(50)")]
        public string? MgCode { get; set; }


        [Column("ORG_CODE", TypeName = "NVARCHAR(50)")]
        public string? OrgCode { get; set; }


        [Column("OPINION_CODE", TypeName = "NVARCHAR(50)")]
        public string? OpinionCode { get; set; }


        [Column("STATUS", TypeName = "NVARCHAR(50)")]
        public string? Status { get; set; }

        [Column("ORG_IN_CHARGE", TypeName = "NVARCHAR(50)")]
        public string? OrgInCharge { get; set; }


        [Column("CONTENT_ORG", TypeName = "NVARCHAR(2500)")]
        public string? ContentOrg { get; set; }

        [Column("CONTENT_REPORT", TypeName = "NVARCHAR(2500)")]
        public string? ContentReport { get; set; }

        [Column("FILE_ID", TypeName = "NVARCHAR(50)")]
        public string? FileId { get; set; }
    }
}
