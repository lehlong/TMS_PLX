using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using DMS.CORE.Common;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_CUSTOMER_PHONE")]
    public class TblMdCustomerPhone : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("CUSTOMER_CODE", TypeName = "VARCHAR(50)")]
        public string CustomerCode { get; set; }

        [Column("PHONE", TypeName = "VARCHAR(20)")]
        public string Phone { get; set; }
    }
}
