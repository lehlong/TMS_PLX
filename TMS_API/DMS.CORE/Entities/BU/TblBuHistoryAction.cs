using DMS.CORE.Common;
using DMS.CORE.Entities.BU;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_HISTORY_ACTION")]
    public class TblBuHistoryAction : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("ACTION", TypeName = "NVARCHAR(255)")]
        public string? Action { get; set; }

        [Column("CONTENTS", TypeName = "NVARCHAR(500)")]
        public string? Contents { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { get; set; }

    }
}
