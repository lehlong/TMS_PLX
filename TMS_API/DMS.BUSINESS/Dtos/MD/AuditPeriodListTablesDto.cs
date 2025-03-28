using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class AuditPeriodListTablesDto : BaseMdTemDto, IMapFrom, IDto
    {
        [Key]
        public int Code { get; set; }
        public string? ListTablesCode { get; set; }
        public string? AuditPeriodCode { get; set; }
        public int? Version { get; set; }
        public string? Status { get; set; }
        public string? TextContent { get; set; }
        public List<AuditTemplateListTablesDataDto>? AuditTemplateListTablesDataReferences { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdAuditPeriodListTables, AuditPeriodListTablesDto>().ReverseMap();
        }
    }
    public class AuditPeriodListTablesCreateDto : BaseMdTemDto, IMapFrom, IDto
    {
       
        public string? ListTablesCode { get; set; }
        public string? AuditPeriodCode { get; set; }
        public int? Version { get; set; }
        public string? Status { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdAuditPeriodListTables, AuditPeriodListTablesCreateDto>().ReverseMap();
        }
    }
}
