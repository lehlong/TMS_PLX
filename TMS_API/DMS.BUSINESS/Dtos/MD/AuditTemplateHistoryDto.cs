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

namespace DMS.BUSINESS.Dtos.MD
{
    public class AuditTemplateHistoryDto : BaseMdTemDto, IMapFrom, IDto
    {
        [Key]
        public string Id { get; set; }
        public string? ListAuditCode { get; set; }
        public string? Action { get; set; }
        public string? TextContent { get; set; }
        public int? AuditPeriodListTablesCode { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdAuditTemplateHistory, AuditTemplateHistoryDto>().ReverseMap();
        }
    }
}
