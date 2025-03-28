using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class MgOpinionListDto : BaseMdDto,IMapFrom,IDto
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }
        [Key]
        [Description("Mã kiến nghị")]
        public string Code { get; set; }
        [Description("Tên kiến nghị")]
        public string Name { get; set; }
        [Description("Ghi chú")]
        public string Description { get; set; }
        [Description("Năm")]
        public string TimeYear { get; set; }
        [Description("Mã đợt")]
        public string AuditPeriod { get; set; }
        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }
        public virtual TblMdAuditPeriod AuditPeriods { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMgOpinionList, MgOpinionListDto>().ReverseMap();
        }
    }
    public class MgOpinionListCreateDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public string Code { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string TimeYear { get; set; }
        public string AuditPeriod { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMgOpinionList, MgOpinionListCreateDto>().ReverseMap();
        }
    
}
}
