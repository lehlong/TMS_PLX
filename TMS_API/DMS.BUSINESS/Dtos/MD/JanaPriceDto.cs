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
    public class JanaPriceDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Mã phương thức phân phối")]
        public string PtppCode { get; set; }

        [Description("Mã mặt hàng")]
        public string GoodsCode { get; set; }

        [Description("Giá nhập mua có Vat")]
        public decimal? GiaNhapMuaVat { get; set; }

        [Description("Lương nhập xuất nội bộ")]
        public decimal? LuongNxNoiBo { get; set; }

        [Description("CHi phí hỗ trợ bán hàng")]
        public decimal? CpHoTro { get; set; }

        [Description("Giá đại lý có Vat")]
        public decimal? GiaDaiLyVat { get; set; }

        [Description("Tiền Lương bán hàng")]
        public decimal? LuongBanHang { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdJanaPrice, JanaPriceDto>().ReverseMap();
        }

    }
}
