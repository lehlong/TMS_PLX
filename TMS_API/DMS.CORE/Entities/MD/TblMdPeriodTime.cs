using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_PERIOD_TIME")]
    public class TblMdPeriodTime : BaseEntity
    {
        [Key]
        [Column("TIME_YEAR", TypeName = "INT")]
        public int TimeYear { get; set; }
        [Column("IS_CLOSED", TypeName ="BIT")]
        public bool IsClosed { get; set; }
      
        [Column("IS_DEFAULT", TypeName ="BIT")]
        public bool IsDefault { get; set; }
    }
}
