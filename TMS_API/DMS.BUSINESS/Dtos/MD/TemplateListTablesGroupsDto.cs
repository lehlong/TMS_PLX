using AutoMapper;
using Common;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.MD
{
    public class TemplateListTablesGroupsDto : BaseMdDto, IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string? Id { get; set; }
        public string Name { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateListTablesGroups, TemplateListTablesGroupsDto>().ReverseMap();
        }
    }
}
