using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.BU;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class TemplateListTablesDto : BaseMdTemDto, IMapFrom, IDto
    {
        [Key]
        public string Code { get; set; }
        public string Name { get; set; }
        public string TimeYear { get; set; }
        public string Note { get; set; }
        public string OrgCode { get; set; }
        public string MgListTablesCode { get; set; }
        public Guid? GroupCode { get; set; }
        public bool? IsChecked { get; set; }
        public int? AuditPeriodCode { get; set; }
        public List<TemplateListTablesDataUpdateDto>? TemDataReferences { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTables, TemplateListTablesDto>().ReverseMap();
        }

    }
    public class TemplateListTablesDetailDto : IMapFrom, IDto
    {
        [Key]
        public string Code { get; set; }
        public string Name { get; set; }
        public string TimeYear { get; set; }
        public string Note { get; set; }
        public string OrgCode { get; set; }
        public string? MgListTablesCode { get; set; }
        public Guid? GroupCode { get; set; }
        public OrganizeDto TreeOrganize { get; set; }
        public ListTablesDto TreeListTables { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTables, TemplateListTablesDetailDto>().ReverseMap();
        }
    }

}
