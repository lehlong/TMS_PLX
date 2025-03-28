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
    [Table("T_BU_LIST_TABLES")] 
    public class TblBuListTables : SoftDeleteEntity
    {
        [Key]
        [Column("CODE")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid Code { get; set; }
        [Column("ID", TypeName = "VARCHAR(50)")]
        public string Id { get; set; }
        [Column("NAME", TypeName = "NVARCHAR(MAX)")]
        public string Name { get; set; }
        [Column("PID", TypeName = "VARCHAR(50)")]
        public string PId { get; set; }
        [Column("ORDER_NUMBER", TypeName = "INT")]
        public int OrderNumber { get; set; }
        [Column("MG_CODE", TypeName = "VARCHAR(50)")]
        public string? MgCode { get; set; }
        [ForeignKey("MgCode")]
        public virtual TblMdMgListTables? MgListTables { get; set; }
        [Column("CURRENCY_CODE", TypeName = "VARCHAR(50)")]
        public string? CurrencyCode { get; set; }
        [ForeignKey("CurrencyCode")]
        public virtual TblMdCurrency? Currency { get; set; }
    }
}
