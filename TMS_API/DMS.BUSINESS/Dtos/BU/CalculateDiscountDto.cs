using AutoMapper;
using Common;
using DMS.CORE.Entities.BU;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.BU
{
    public class CalculateDiscountDto : BaseMdDto, IMapFrom, IDto
    {
        
        [Key]
        public string Id { get; set; }
        public string Name { get; set; }
        public DateTime Date { get; set; }
        public string? Status { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblBuCalculateDiscount, CalculateDiscountDto>().ReverseMap();
        }
    }
}
