using AutoMapper;
using Common;
using DMS.CORE.Entities.AD;

using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.AD
{
    public class EmailConfigDto : IDto, IMapFrom
    {
        [Description("ID")]
        public int Id { get; set; }

        [Required]
        [Description("Cổng")]
        public int Port { get; set; }

        [Required]
        [Description("Máy chủ")]
        public string Host { get; set; }

        [Description("Mật khẩu")]
        public string? Pass { get; set; }

        [Required]
        [Description("Địa chỉ email")]
        [EmailAddress]
        public string Email { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblEmailConfig, EmailConfigDto>().ReverseMap();
        }
    }
}
