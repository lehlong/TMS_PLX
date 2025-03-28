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
    [Table("T_BU_HISTORY_DOWNLOAD")]
    public class TblBuHistoryDownload : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "NVARCHAR(500)")]
        public string Code { get; set; }
        [Column("HEADER_CODE", TypeName = "NVARCHAR(500)")]
        public string HeaderCode { get; set; }

        [Column("NAME", TypeName = "NVARCHAR(500)")]
        public string Name { get; set; }

        [Column("TYPE", TypeName = "NVARCHAR(500)")]
        public string? Type { get; set; }

        [Column("PATH", TypeName = "NVARCHAR(MAX)")]
        public string? Path { get; set; }


    }
}
