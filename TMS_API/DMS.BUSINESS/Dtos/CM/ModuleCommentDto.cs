using AutoMapper;
using Common;
using DMS.CORE.Entities.BU;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.BU
{
    public class ModuleCommentDto : IMapFrom, IDto
    {
        [Key]
        public int Id { get; set; }

        public Guid ReferenceId { get; set; }

        public int CommentId { get; set; }

        public string ModuleType { get; set; }

        public virtual ModuleCommentDto Comment { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblCmModuleComment, ModuleCommentDto>().ReverseMap();
        }
    }
}
