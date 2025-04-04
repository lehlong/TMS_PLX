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
    [Table("T_IN_MARKET_COMPETITOR")]
    public class TblInMarketCompetitor : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "NVARCHAR(100)")]
        public string Code { get; set; }
        
        [Column("HEADER_CODE", TypeName = "NVARCHAR(100)")]
        public string HeaderCode { get; set; }

        [Column("GAP", TypeName = "DECIMAL(18,0)")]
        public decimal? Gap { get; set; }

        [Column("MARKET_CODE", TypeName = "NVARCHAR(100)")]
        public string? MarketCode { get; set; }

        [Column("MARKET_NAME", TypeName = "NVARCHAR(100)")]
        public string? MarketName { get; set; }

        [Column("COMPETITOR_CODE", TypeName = "NVARCHAR(100)")]
        public string? CompetitorCode { get; set; }

        [Column("COMPETITOR_NAME", TypeName = "NVARCHAR(100)")]
        public string? CompetitorName { get; set; }


    }
}
