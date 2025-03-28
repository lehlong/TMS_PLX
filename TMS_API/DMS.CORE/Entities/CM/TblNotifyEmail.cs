using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DMS.CORE.Entities.AD;

namespace DMS.CORE.Entities.BU
{
    [Table("T_CM_NOTIFY_EMAIL")]
    public class TblNotifyEmail : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public string Id{ get; set; }
        [Column("HEADER_ID")]
        public string HeaderId { get; set; }

        [Column("STATUS")]
        public int? Status { get; set; }

        [Column("SUBJECT")]
        public string Subject { get; set; }

        [Column("CONTENTS")]
        public string Contents { get; set; }

        [Column("NUMBER_RETRY")]
            public int? NumberRetry { get; set; }
        [Column("IS_SEND")]
        public string IsSend { get; set; }
        [Column("EMAIL")]
        public string Email { get; set; }




    }
}
