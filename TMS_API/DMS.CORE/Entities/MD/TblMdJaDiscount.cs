using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.MD
{
    [Table("T_MD_JA_DISCOUNT")]
    public class TblMdJaDiscount : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("SAN_LUONG_CODE", TypeName = "NVARCHAR(255)")]
        public string SanLuongCode { get; set; }

        [Column("GOODS_CODE", TypeName = "NVARCHAR(255)")]
        public string GoodsCode { get; set; }

        [Column("AREA_CODE", TypeName = "NVARCHAR(255)")]
        public string AreaCode { get; set; }

        [Column("CHIET_KHAU", TypeName = "DECIMAL(18,5)")]
        public decimal ChietKhau { get; set; }

    }
}
