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
    public class GoodsDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Tên")]
        public string Name { get; set; }

        [Description("Kiểu mặt hàng")]
        public string Type { get; set; }

        [Description("Thuế bảo vệ môi trường")]
        public decimal ThueBvmt { get; set; }

        [Description("Hệ số VFC BQ mùa miền")]
        public decimal Vfc { get; set; }

        [Description("Mức Tăng So với V1")]
        public decimal? MtsV1 { get; set; }
        public int Order { get; set; }

        [Description("Ngày tạo")]
        public DateTime createDate { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdGoods, GoodsDto>().ReverseMap();
        }

    }
}
