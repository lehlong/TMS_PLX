using Common;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using AutoMapper;
using DMS.CORE.Entities.MD;

namespace DMS.BUSINESS.Dtos.MD
{
    public class CuocVanChuyenListDto : BaseMdDto, IMapFrom, IDto
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Mã")]
        public string Code { set; get; }

        [Description("Ngày tạo")]
        public DateTime createDate { get; set; }

        [Description("Tên đợt nhập cước vận chuyển")]
        public string Name { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdCuocVanChuyenList, CuocVanChuyenListDto>().ReverseMap();
        }
    }
}
