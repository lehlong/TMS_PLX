using AutoMapper;
using Common;
using DMS.CORE.Entities.BU;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.BU
{
    public class CalculateResultListDto : BaseMdDto, IMapFrom, IDto
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }

        [Key]
        [Description("Code")]
        public string Code { get; set; }

        [Description("Tên")]
        public string Name { get; set; }

        [Description("Ngày bắt đầu")]
        public DateTime FDate { get; set; }

        [Description("Mã người ký")]
        public string SignerCode { get; set; }

        [Description("Quyết định số")]
        public string QuyetDinhSo { get; set; }

        [Description("Status")]
        public string? Status { get; set; }

        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblBuCalculateResultList, CalculateResultListDto>().ReverseMap();
        }
    }
}
