using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.BU;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_OPINION_LIST_ORGANIZE")]
    public class TblMdOpinionListOrganize
    {
        [Column("OPINION_LIST_CODE")]
        public Guid OpinionListCode { get; set; }

        [ForeignKey("OpinionListCode")]
        public virtual tblBuOpinionList OpinionList { get; set; }

        [Column("ORGANIZE_ID")]
        public string OrganizeId { get; set; }

        [ForeignKey("OrganizeId")]
        public virtual tblAdOrganize Organize { get; set; }
        [Column("IS_PENDING")]
        public bool IsPending { get; set; }
    }
}
