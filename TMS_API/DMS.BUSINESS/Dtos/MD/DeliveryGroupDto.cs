using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using AutoMapper;
using DMS.CORE.Entities.MD;
using Common;

namespace DMS.BUSINESS.Dtos.MD
{
    public class DeliveryGroupDto : BaseMdDto, IDto, IMapFrom
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { get; set; }

        [Description("Tên")]
        public string Name { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdDeliveryGroup, DeliveryGroupDto>().ReverseMap();
        }

    }
}
