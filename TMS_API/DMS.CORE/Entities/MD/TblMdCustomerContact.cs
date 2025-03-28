using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_CUSTOMER_CONTACT")]
    public class TblMdCustomerContact : BaseEntity
    {

            [Key]
            [Column("CODE", TypeName = "VARCHAR(50)")]
            public string? Code { set; get; }

            [Column("CUSTOMER_CODE", TypeName = "VARCHAR(50)")]
            public string? Customer_Code { set; get; }

            [Column("TYPE", TypeName = "VARCHAR(10)")]
            public string Type { set; get; }

            [Column("VALUE", TypeName = "NVARCHAR(250)")]
            public string? Value { set; get; }

        }
    
}
