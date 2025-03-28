using AutoMapper;
using Common;
using DMS.BUSINESS.Common.Enum;
using DMS.BUSINESS.Dtos.BU;
using DMS.COMMON.Common.Enum;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class ListAuditDto : BaseMdTemDto, IMapFrom, IDto
    {
        [Key]
        public string? Code { get; set; }
        public string? Name { get; set; }
        public string? TimeYear { get; set; }
        public string? AuditPeriod { get; set; }
        public DateTime? ReportDate { get; set; }
        public string? ReportNumber { get; set; }
        public string? Status { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string? OpinionCode { get; set; }
        public string? Note { get; set; }
        public Guid? FileId { get; set; }
        public string? TextContent { get; set; }
        public string? Approver { get; set; }

        public string State { get => this.IsActive == true ? "Đang hoạt động" : "Khóa"; }
        public List<tblMdListAuditHistory>? History { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdListAudit, ListAuditDto>().ReverseMap();
        }
    }
    public class OpinionStatistic
    {
        public string AuditCode { get; set; }
        public string OrgCode { get; set; }
        public string Opinion { get; set; }
        public string Status { get; set; }
    }

}
