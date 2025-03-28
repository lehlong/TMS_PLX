using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.CORE.Entities.AD;
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
    public class MgListTablesGroupsDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string? Id { get; set; }
        public string Name { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdMgListTablesGroups, MgListTablesGroupsDto>().ReverseMap();
        }
    }
}

