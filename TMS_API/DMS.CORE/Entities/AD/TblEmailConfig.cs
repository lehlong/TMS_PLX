using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.AD
{
    [Table("EMAIL_CONFIG")]
    public class TblEmailConfig : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public int Id { get; set; }

        [Column("PORT")]
        public int Port { get; set; }

        [Column("HOST", TypeName = "NVARCHAR(100)")]
        public string Host { get; set; }

        [Column("PASS", TypeName = "NVARCHAR(100)")]
        public string? Pass { get; set; }

        [Column("EMAIL", TypeName = "NVARCHAR(255)")]
        public string Email { get; set; }
    }
}
