using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_PENDING_OPINION_MAPPING")]
    public class tblBuPendingOpinionMapping : SoftDeleteEntity
    {
        [Key]
        [Column("ID")]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public Guid Code { get; set; }
        [Column("UNFINISHED_ID", TypeName = "VARCHAR(50)")]
        public string Unfinished_Id { get; set; }
        [Column("PENDING_ID", TypeName = "VARCHAR(50)")]
        public string Pending_Id { get; set; }

    }
}
