using System.ComponentModel.DataAnnotations;
using AutoMapper;
using System.Text.Json.Serialization;
using System.ComponentModel;

using Common;
using DMS.CORE.Entities.IN;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.BUSINESS.Dtos.IN
{
    public class CustomerFobDto : BaseMdDto, IMapFrom, IDto
    {
        [JsonIgnore]
        [Description("Số thứ tự")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string? Code { get; set; }

        [Description("Mã Header")]
        public string? HeaderCode { get; set; }

        [Description("Mã khách hàng")]
        public string? CustomerCode { get; set; }

        [Description("Tên khách hàng")]
        public string? CustomerName { get; set; }

        [Description("Người đại diện")]
        public string? CustomerFob { get; set; }

        [Description("Mức giảm giá tại kho bên bán")]
        public string? Fob { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TbLInCustomerFob, CustomerFobDto>().ReverseMap();
        }
    }
}
