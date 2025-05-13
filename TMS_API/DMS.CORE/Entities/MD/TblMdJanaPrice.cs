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
    [Table("T_MD_JANA_PRICE")]
    public class TblMdJanaPrice : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("PTPP_CODE", TypeName = "VARCHAR(50)")]
        public string PtppCode { get; set; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]
        public string GoodsCode { get; set; }

        [Column("GIA_NHAP_MUA_VAT", TypeName = "DECIMAL(18,0)")]
        public decimal GiaNhapMuaVat { get; set; }

        [Column("LUONG_NX_NOI_BO", TypeName = "DECIMAL(18,0)")]
        public decimal LuongNxNoiBo { get; set; }

        [Column("CP_HO_TRO", TypeName = "DECIMAL(18,0)")]
        public decimal CpHoTro { get; set; }

        [Column("GIA_DAI_LY_VAT", TypeName = "DECIMAL(18,0)")]
        public decimal GiaDaiLyVat { get; set; }

        [Column("LUONG_BAN_HANG", TypeName = "DECIMAL(18,0)")]
        public decimal LuongBanHang { get; set; }

    }
}
