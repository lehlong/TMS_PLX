using Common;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AutoMapper;
using DMS.CORE.Entities.MD;
using System.ComponentModel.DataAnnotations.Schema;

namespace DMS.BUSINESS.Dtos.MD
{
    public class TemplateReportElementDto : BaseMdDto, IDto, IMapFrom
    {
        [Key]
        public string Id { get; set; }
        public string FileId { get; set; }
        public string TextElement { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateReportElement, TemplateReportElementDto>().ReverseMap();
        }
    }
}
