using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.IN
{
    [Table("T_IN_CUSTOMER_FOB")]
    public class TbLInCustomerFob : BaseEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string? Code { get; set; }

        [Column("HEADER_CODE", TypeName = "VARCHAR(50)")]
        public string? HeaderCode { set; get; }

        [Column("CUSTOMER_NAME", TypeName = "NVARCHAR(255)")]
        public string? CustomerName { set; get; }

        [Column("CUSTOMER_CODE", TypeName = "VARCHAR(50)")]
        public string? CustomerCode { get; set; }

        [Column("FOB", TypeName = "DECIMAL(18, 0)")]
        public decimal? Fob { get; set; }
    }
}
