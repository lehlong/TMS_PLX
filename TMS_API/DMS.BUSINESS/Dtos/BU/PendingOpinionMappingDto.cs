using AutoMapper;
using Common;
using DMS.CORE.Entities.BU;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.BU
{
    public class PendingOpinionMappingDto : IMapFrom, IDto
    {
        [Key]
        public Guid Id { get; set; }
        public string Unfinished_Id { get; set; }
        public string Pending_Id { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<tblBuPendingOpinionMapping, PendingOpinionMappingDto>().ReverseMap();
        }
    }
}
