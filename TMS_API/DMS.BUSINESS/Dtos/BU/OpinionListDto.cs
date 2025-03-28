using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics.Contracts;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.BU
{
    public class OpinionListDto : IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string PId { get; set; }
        public int? OrderNumber { get; set; }
        public bool Checked { get; set; }   
        public string? Title { get; set; }
        public string? Key { get; set; }
        public string? Account { get; set; }
        public string? MgCode { get; set; }
        public bool Expanded { get; set; } = true;
        public List<OpinionListOrganizeDto>? OrganizeReferences { get; set; }
        public List<OpinionListDto>? Children { get; set; }
        public List<OrganizeDto>? PendingOrganizes { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<tblBuOpinionList, OpinionListDto>().ReverseMap();
        }
    }
    public class OpinionViewModel
    {
        public Guid GroupId { get; set; }

        public List<OpinionListDto> ListOpinion { get; set; }

        public OpinionViewModel()
        {
            ListOpinion = new List<OpinionListDto>();
        }
    }
    public class OpinionListUpdateDto : IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string PId { get; set; }
        public int? OrderNumber { get; set; }
        public bool IsChecked { get; set; }
        public string? Title { get; set; }
        public string? Key { get; set; }
        public string? Account { get; set; }
        public string? MgCode { get; set; }
        public List<OpinionListOrganizeUpdateDto>? OrganizeReferences { get; set; }
        public List<OpinionListDto>? Children { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<tblBuOpinionList, OpinionListUpdateDto>().ReverseMap();
        }
    }

    public class OpinionListDetailDto : IMapFrom, IDto
    {
        [Key]
        public Guid Code { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string PId { get; set; }
        public int? OrderNumber { get; set; }
        public bool IsChecked { get; set; }
        public string? Title { get; set; }
        public string? Key { get; set; }
        public string? Account { get; set; }
        public string? MgCode { get; set; }
        public OrganizeDto TreeOrganize { get; set; }
        public List<OpinionListDto>? Children { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<tblBuOpinionList, OpinionListDetailDto>().ReverseMap();
        }
    }

    //public class OpinionListCreateDto : IMapFrom, IDto
    //{
    //    [Key]
    //    public Guid Code { get; set; }
    //    public string Id { get; set; }
    //    public string Name { get; set; }
    //    public string PId { get; set; }
    //    public int? OrderNumber { get; set; }

    //    public bool IsChecked { get; set; }

    //    public string? Title { get; set; }

    //    public string? Key { get; set; }
    //    public string? MgCode { get; set; }
    //    public List<OpinionListOrganizeDto>? OrganizeReferences { get; set; }
    //    public List<OpinionListDto>? Children { get; set; }
    //    public void Mapping(Profile profile)
    //    {
    //        profile.CreateMap<tblBuOpinionList, OpinionListDto>().ReverseMap();
    //    }
    //}

}
