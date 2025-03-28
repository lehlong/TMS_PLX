using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class MapPointCustomerGoodsDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }
        [Key]
        [Description("Mã ")]
        public string Code { get; set; }
        
        [Description("Mã điểm nhận hàng")]
        public string? DeliveryPointCode { get; set; }

        [Description("Mã khách hàng")]
        public string? CustomerCode { get; set; }

        [Description("Mã Mặt hàng")]
        public string? GoodsCode { get; set; }

        [Description("Cước vận chuyển bình quân")]
        public decimal? CuocVcBq { get; set; }

        [Description("Kiểu ")]
        public string? Type { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMapPointCustomerGoods, MapPointCustomerGoodsDto>().ReverseMap();
        }

    }
}
