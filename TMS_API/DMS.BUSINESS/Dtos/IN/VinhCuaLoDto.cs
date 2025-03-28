using System.ComponentModel.DataAnnotations;
using AutoMapper;
using System.Text.Json.Serialization;
using System.ComponentModel;

using Common;
using DMS.CORE.Entities.IN;

namespace DMS.BUSINESS.Dtos.IN
{
    public class VinhCuaLoDto : BaseMdDto, IMapFrom, IDto
    {
        [JsonIgnore]
        [Description("Số thứ tự")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Mã mặt hàng")]
        public string GoodsCode { get; set; }

        [Description("Mã Header")]
        public string HeaderCode { get; set; }

        [Description("Giá bán lẻ cơ sở v1")]
        public decimal GblcsV1 { get; set; }

        [Description("Giá bán lẻ v2")]
        public decimal GblV2 { get; set; }

        [Description("Chênh lệch giá V2-V1")]
        public decimal V2_V1 { get; set; }

        [Description("Mức tăng so với V1")]
        public decimal MtsV1 { get; set; }

        [Description("Giá niêm yết")]
        public decimal Gny { get; set; }

        [Description("Chênh lệch giá bán lẻ Vùng trung tâm và còn lại")]
        public decimal Clgblv { get; set; }

        [Description("Từ ngày")]
        public DateTime? FromDate { get; set; }

        [Description("Đến ngày")]
        public DateTime? ToDate { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblInVinhCuaLo, VinhCuaLoDto>().ReverseMap();
        }
    }

}
