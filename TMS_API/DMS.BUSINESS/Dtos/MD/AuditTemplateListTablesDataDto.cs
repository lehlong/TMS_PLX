using AutoMapper;
using Common;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class AuditTemplateListTablesDataDto : IMapFrom, IDto
    {
        public Guid TemplateDataCode { get; set; }

        public int AuditListTablesCode { get; set; }
        public string? Unit { get; set; }

        public decimal? AuditValue { get; set; }

        public string? AuditExplanation { get; set; }
       
        public decimal? ExplanationValue { get; set; }

        public string? ExplanationNote { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdAuditTemplateListTablesData, AuditTemplateListTablesDataDto>().ReverseMap();
        }
    }
    public class AuditTemplateListTablesDataListDto : IMapFrom, IDto
    {
        public Guid TemplateDataCode { get; set; }

        public int AuditListTablesCode { get; set; }
        public Guid? ListTablesCode { get; set; }

        [Description("Mã đơn vị")]
        public string? OrgCode { get; set; }
        [Description("Tên đơn vị")]
        public string? OrgName { get; set; }
        [Description("Mã chỉ tiêu")]
        public string? ListTablesId { get; set; }
       
        [Description("Tên chỉ tiêu")]
        public string? ListTablesName { get; set; }
        public string? TemplateCode { get; set; }

        [Description("Đơn vị")]
        public string? Unit { get; set; }

        [Description("Giá trị kiểm toán")]
        public decimal? AuditValue { get; set; }

        [Description("Thuyết minh của kiểm toán")]
        public string? AuditExplanation { get; set; }
        [Description("Giá trị giải trình")]
        public decimal? ExplanationValue { get; set; }
        [Description("Thuyết minh giải trình")]
        public string? ExplanationNote { get; set; }
        public int? OrderNumber { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdAuditTemplateListTablesData, AuditTemplateListTablesDataListDto>().ReverseMap();
        }
    }

}
