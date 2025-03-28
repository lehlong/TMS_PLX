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
    [Table("T_MD_MAP_POINT_CUSTOMER_GOODS")]
    public class TblMdMapPointCustomerGoods : SoftDeleteEntity
    {
        [Key]
        [Column("CODE", TypeName = "VARCHAR(50)")]
        public string Code { get; set; }
        
        [Column("DELIVERY_POINT_CODE", TypeName = "VARCHAR(50)")]
        public string DeliveryPointCode { get; set; }

        [Column("CUSTOMER_CODE", TypeName = "VARCHAR(50)")]
        public string CustomerCode { get; set; }

        [Column("GOODS_CODE", TypeName = "VARCHAR(50)")]
        public string GoodsCode { get; set; }

        [Column("CUOC_VC_BQ", TypeName = "DECIMAL(18,0)")]
        public decimal? CuocVcBq { set; get; }

        [Column("TYPE", TypeName = "VARCHAR(50)")]
        public string? Type { get; set; }

    }
}
