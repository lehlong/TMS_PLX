using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DMS.CORE.Common;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_RETAIL_PRICE_LIST")]

    public class TblMdRetailPriceList : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(1000)")]
        public string Code { set; get; }

        [Column("NAME", TypeName = "NVARCHAR(1000)")]
        public string? Name { set; get; }

        [Column("F_DATE", TypeName = "DATETIME")]
        public DateTime? FDate { set; get; }

        [Column("T_DATE", TypeName = "DATETIME")]
        public DateTime ToDate { set; get; }
    }
}
