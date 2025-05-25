using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.CORE.Entities.AD
{
    [Table("SMS_CONFIG")]
    public class TblSmsConfig : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public int Id { get; set; }

        [Column("URLSMS", TypeName = "NVARCHAR(255)")]
        public string UrlSms { get; set; }

        [Column("USERNAME", TypeName = "NVARCHAR(100)")]
        public string Username { get; set; }

        [Column("PASSWORD", TypeName = "NVARCHAR(100)")]
        public string Password { get; set; }

        [Column("CPCODE", TypeName = "NVARCHAR(100)")]
        public string CpCode { get; set; }

        [Column("SERVICEID", TypeName = "NVARCHAR(100)")]
        public string ServiceId { get; set; }
    }
}
