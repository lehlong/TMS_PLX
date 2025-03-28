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
    public class GiaGiaoTapDoanDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Mã mặt hàng")]
        public string GoodsCode { get; set; }

        [Description("Mã khách hàng")]
        public string GgtdlCode { get; set; }

        [Description("Giá cũ")]
        public decimal? OldPrice { get; set; }

        [Description("Giá mới")]
        public decimal? NewPrice { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdGiaGiaoTapDoan, GiaGiaoTapDoanDto>().ReverseMap();
        }
    }
}
