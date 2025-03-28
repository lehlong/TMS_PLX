using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_CALCULATE_DISCOUNT")]
    public class TblBuCalculateDiscount : SoftDeleteEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }

        [Column("NAME", TypeName = "NVARCHAR(500)")]
        public string Name { get; set; }

        [Column("DATE", TypeName = "DATETIME")]
        public DateTime Date { get; set; }

        [Column("STATUS", TypeName = "NVARCHAR(50)")]
        public string? Status { get; set; }
    }
}
