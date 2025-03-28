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
    public class TemplateReportDto : BaseMdDto, IDto, IMapFrom
    {

        [Key]
        public string Id { get; set; }
        public string Year { get; set; }
        public string AuditPeriod { get; set; }
        public string FileName { get; set; }
        public string FileOldName { get; set; }
        public string FileExt { get; set; }
        public string FileSize { get; set; }
        public string NetworkPath { get; set; }
        public string FullPath { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblMdTemplateReport, TemplateReportDto>().ReverseMap();
        }
    }
}
