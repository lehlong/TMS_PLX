using System.ComponentModel.DataAnnotations;
using AutoMapper;
using System.Text.Json.Serialization;
using System.ComponentModel;

using Common;
using DMS.CORE.Entities.IN;

namespace DMS.BUSINESS.Dtos.IN
{
    public class NguoiKyTenDto : BaseMdDto, IMapFrom, IDto
    {
        [JsonIgnore]
        [Description("Số thứ tự")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string? Code { get; set; }

        [Description("Mã Header")]
        public string? HeaderCode { get; set; }

        [Description("Đại diện")]
        public string? DaiDien { get; set; }

        [Description("Người đại diện")]
        public string? NguoiDaiDien { get; set; }

        [Description("Quyết định số")]
        public string? QuyetDinhSo { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblInNguoiKyTen, NguoiKyTenDto>().ReverseMap();
        }
    }

}
