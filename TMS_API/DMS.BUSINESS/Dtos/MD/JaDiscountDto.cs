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
    public class JaDiscountDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã ")]
        public string Code { get; set; }

        [Description("Sản lượng")]
        public string SanLuongCode { get; set; }

        [Description("Sản phẩm")]
        public string GoodsCode { get; set; }

        [Description("AREA_CODE")]
        public string AreaCode { get; set; }

        [Description("Mức chiết khấu")]
        public decimal ChietKhau { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdJaDiscount, JaDiscountDto>().ReverseMap();
        }

    }
}
