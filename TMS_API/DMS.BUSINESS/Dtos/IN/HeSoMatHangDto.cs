using System.ComponentModel.DataAnnotations;
using AutoMapper;
using System.Text.Json.Serialization;
using System.ComponentModel;

using Common;
using DMS.CORE.Entities.IN;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.BUSINESS.Dtos.IN
{
    public class HeSoMatHangDto : BaseMdDto, IMapFrom, IDto
    {
        [JsonIgnore]
        [Description("Số thứ tự")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string? Code { get; set; }

        [Description("Mã mặt hàng")]
        public string GoodsCode { get; set; }

        [Description("Mã Header")]
        public string HeaderCode { get; set; }

        [Description("Hệ số VCF BQ Mùa miền")]
        public decimal HeSoVcf { get; set; }

        [Description("Thuế BVMT")]
        public decimal ThueBvmt { get; set; }

        [Description("L15 chưa VAT và BVMT (PT bán lẻ - V2) ")]
        public decimal L15ChuaVatBvmt { get; set; }

        [Description("L15 chưa VAT và BVMT (ngoài bán lẻ)")]
        public decimal? L15ChuaVatBvmtNbl { get; set; }

        [Description("Giảm giá FOB có VAT & BVMT")]
        public decimal GiamGiaFob { get; set; }

        [Description("Lãi Gộp điều tiết")]
        public decimal? LaiGopDieuTiet { get; set; }

        [Description("Từ ngày")]
        public DateTime? FromDate { get; set; }

        [Description("Đến ngày")]
        public DateTime? ToDate { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblInHeSoMatHang, HeSoMatHangDto>().ReverseMap();
        }
    }

}
