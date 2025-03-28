using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class OpinionListOrganizeDto : IMapFrom, IDto
    {
        public string OpinionListCode { get; set; }
        public string OrganizeId { get; set; }
        public bool IsPending { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdOpinionListOrganize, OpinionListOrganizeDto>().ReverseMap();
        }
    }
    public class OpinionListOrganizeUpdateDto : IMapFrom, IDto
    {
        public string OrganizeId { get; set; }
        public bool IsPending { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdOpinionListOrganize, OpinionListOrganizeUpdateDto>().ReverseMap();
        }
    }

}
