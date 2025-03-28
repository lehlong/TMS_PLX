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
    [Table("T_MD_GIA_GIAO_TAP_DOAN_LIST")]
    public class TblMdGiaGiaoTapDoanList : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(1000)")]
        public string Code { set; get; }

        [Column("Name", TypeName = "NVARCHAR(1000)")]
        public string? Name { set; get; }

        [Column("F_DATE", TypeName = "DATETIME")]
        public DateTime? FDate { set; get; }

    }
}
