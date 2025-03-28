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
    [Table("T_MD_SALES_METHOD")]

    public class TblMdSalesMethod : SoftDeleteEntity
    {
            [Key]
            [Column("CODE", TypeName = "VARCHAR(50)")]
            public string Code { get; set; }
            [Column("NAME", TypeName = "NVARCHAR(255)")]
            public string Name { get; set; }
    }
}
