using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
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
    public class TemplateListTablesDataListDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        [Description("Mã đơn vị")]
        public string? OrgCode { get; set; }

        [Description("Tên đơn vị")]
        public string? OrgName { get; set; }

        [Description("Mã chỉ tiêu")]
        public string? ListTablesId { get; set; }

        [Description("Tên chỉ tiêu")]
        public string? ListTablesName { get; set; }

        [Description("Đơn vị")]
        public string? Unit { get; set; }

        [Description("Giá trị kiểm toán")]
        public decimal? AuditValue { get; set; }

        [Description("Thuyết minh của kiểm toán")]
        public string? AuditNotes { get; set; }

        [Description("Giá trị giải trình")]
        public decimal? ExplanationValue { get; set; }

        [Description("Thuyết minh giải trình")]
        public string? Explanation { get; set; }
        public string TemplateCode { get; set; }
        public int? OrderNumber { get; set; }
        public bool? IsParent { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTablesData, TemplateListTablesDataListDto>().ReverseMap();
        }
    }

    public class TemplateListTablesDataDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string OrgCode { get; set; }
        public Guid? ListTablesCode { get; set; }
        public string TemplateCode { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTablesData, TemplateListTablesDataDto>().ReverseMap();
        }
    }
    public class TemplateListTablesDataUpdateDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string OrgCode { get; set; }
        public Guid ListTablesCode { get; set; }
        public string TemplateCode { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTablesData, TemplateListTablesDataUpdateDto>().ReverseMap();
        }
    }
    public class TemplateListTablesDataGenCodeDto : IMapFrom, IDto
    {

        public string OrgCode { get; set; }
        public Guid ListTablesCode { get; set; }
        public string TemplateCode { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTablesData, TemplateListTablesDataGenCodeDto>().ReverseMap();
        }
    }


}
