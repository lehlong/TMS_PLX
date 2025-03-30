using DMS.CORE.Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Entities.BU
{
    [Table("T_BU_INPUT_MARKET")]
    public class TblBuInputMarket : SoftDeleteEntity
    {
        [Key]
        [Column("ID", TypeName = "NVARCHAR(50)")]
        public string Id { get; set; }
        [Column("HEADER_ID", TypeName = "NVARCHAR(50)")]
        public string HeaderId { get; set; }

        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }

        [Column("NAME", TypeName = "NVARCHAR(255)")]
        public string Name { get; set; }
        [Column("FULL_NAME", TypeName = "NVARCHAR(500)")]
        public string? FullName { get; set; }

        [Column("LOCAL2", TypeName = "NVARCHAR(500)")]
        public string? Local2 { get; set; }

        [Column("GAP", TypeName = "DECIMAL(18,0)")]
        public decimal? Gap { get; set; }

        [Column("COEFFICIENT", TypeName = "DECIMAL(18,1)")]
        public decimal? Coefficient { get; set; }

        [Column("LOCAL_CODE", TypeName = "VARCHAR(50)")]
        public string LocalCode { get; set; }

        [Column("WAREHOUSE_CODE", TypeName = "VARCHAR(50)")]
        public string? WarehouseCode { get; set; }

        [Column("CUOC_VC_BQ", TypeName = "DECIMAL(18,0)")]
        public decimal? CuocVCBQ { get; set; }

        [Column("CP_CHUNG_CHUA_CUOC_VC", TypeName = "DECIMAL(18,0)")]
        public decimal? CPChungChuaCuocVC { get; set; }

        [Column("CK_DIEU_TIET_XANG", TypeName = "DECIMAL(18,0)")]
        public decimal? CkDieuTietXang { get; set; }

        [Column("CK_DIEU_TIET_DAU", TypeName = "DECIMAL(18,0)")]
        public decimal? CkDieuTietDau { get; set; }

    }
}
