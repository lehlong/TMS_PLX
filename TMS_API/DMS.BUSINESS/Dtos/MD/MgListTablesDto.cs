using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using DMS.CORE.Entities.MD;
using DMS.BUSINESS.Dtos.AD;
using System.ComponentModel;

namespace DMS.BUSINESS.Dtos.MD
{
    public class MgListTablesDto : BaseMdDto, IMapFrom, IDto
    {
        [Description("STT")]
        public int OrdinalNumber { get; set; }
        [Key]
        [Description("Mã bảng biểu")]
        public string Code { get; set; }
        [Description("Tên bảng biểu")]
        public string Name { get; set; }
        [Description("Ghi chú")]
        public string Description { get; set; }
        [Description("Năm")]
        public string TimeYear { get; set; }
        [Description("Mã đợt")]
        public string AuditPeriod { get; set; }
        public Guid? GroupCode { get; set; }
        [Description("Trạng thái")]
        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }
        public bool? IsChecked { get; set; }
        public string? AuditPeriodCode { get; set; }
        public virtual TblMdAuditPeriod AuditPeriods { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMgListTables, MgListTablesDto>().ReverseMap();
        }
    }
    public class MgListTablesCreateDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public string Code { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string TimeYear { get; set; }
        public string AuditPeriod { get; set; }
        public Guid? GroupCode { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMgListTables, MgListTablesCreateDto>().ReverseMap();
        }
    }

}
