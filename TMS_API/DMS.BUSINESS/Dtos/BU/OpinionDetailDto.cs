using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.BU;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.BU
{
    public class OpinionDetailDto : IMapFrom, IDto
    {
        [Key]
        public string Id { get; set; }
        public string? MgCode { get; set; }
        public string? OrgCode { get; set; }
        public string? OpinionCode { get; set; }
        public string? Status { get; set; }
        public string? ContentOrg { get; set; }
        public string? ContentReport { get; set; }
        public string? OrgInCharge { get; set; }
        public string? CreateBy { get; set; }
        public string? FileId { get; set; }
        public string? Action { get; set; }
        public string? TextContent { get; set; }
        public List<tblBuOpinionDetailHistory> History { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<tblBuOpinionDetail, OpinionDetailDto>().ReverseMap();
        }
    }
}
