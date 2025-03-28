using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.CORE.Entities.BU;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.BU
{
    public class EmailDto : IMapFrom, IDto
    {
        [Key]
        public string Id { get; set; }

        public string? Status { get; set; }

        public string Email { get; set; }

        public string Subject { get; set; }

        public string  Contents { get; set; }
        public string Is_send { get; set; }
        public int? NumberRerey { get; set; }
        public string HeaderId { get; set; }


        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblNotifyEmail, EmailDto>().ReverseMap();
        }
    }


}
