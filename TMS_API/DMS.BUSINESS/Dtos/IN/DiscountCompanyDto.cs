using System.ComponentModel.DataAnnotations;
using AutoMapper;
using System.Text.Json.Serialization;
using System.ComponentModel;

using Common;
using DMS.CORE.Entities.IN;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.BUSINESS.Dtos.IN
{
    public class DiscountCompanyDto
    {
        [JsonIgnore]
        [Description("Số thứ tự")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string? Code { get; set; }

        [Description("Mã Header")]
        public string? HeaderCode { get; set; }

        [Description("Mã Hàng hoá")]
        public string? GoodsCode { get; set; }

        [Description("Mức giảm giá công ty")]
        public string? Discount { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblInDiscountCompany, DiscountCompanyDto>().ReverseMap();
        }
    }
}
