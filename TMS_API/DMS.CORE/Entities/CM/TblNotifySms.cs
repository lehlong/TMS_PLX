using DMS.CORE.Common;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DMS.CORE.Entities.AD;

namespace DMS.CORE.Entities.BU
{
    [Table("T_CM_NOTIFY_SMS")]
    public class TblNotifySms : BaseEntity
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        [Column("ID")]
        public string? Id{ get; set; }

        [Column("HEADER_ID")]
        public string HeaderId { get; set; }

        [Column("STATUS")]
        public string? Status { get; set; }

        [Column("SUBJECT")]
        public string? Subject { get; set; }

        [Column("CONTENTS")]
        public string Contents { get; set; }

        [Column("NUMBER_RETRY")]
        public int? NumberRetry { get; set; }

        [Column("IS_SEND")]
        public string? IsSend { get; set; }

        [Column("PHONE_NUMBER")]
        public string? PhoneNumber { get; set; }

        [Column("CUSTOMER_CODE")]
        public string? CustomerCode { get; set; }

        [Column("MARKET_CODE")]
        public string? MarketCode { get; set; }


    }
    public class NotifyEmailViewModel
    {
        public string Id { get; set; }
        public string HeaderID { get; set; }
        public string Status { get; set; }
        public string Subject { get; set; }
        public string Contents { get; set; }
        public int? NumberRetry { get; set; }
        public string IsSend { get; set; }
        public string Email { get; set; }
        public string CustomerCode { get; set; }
        public string CustomerName { get; set; }
        public string? GroupMailCode { get; set; }
        public string CheckFile { get; set; }
    }
}
