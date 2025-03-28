using AutoMapper;
using Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.CORE.Entities.BU;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.BU
{
    public class CommentDto : IMapFrom, IDto
    {
        [Key]
        public int Id { get; set; }

        public int? PId { get; set; }

        public string Type { get; set; }

        public string Content { get; set; }

        public Guid? AttachmentId { get; set; }

        public virtual List<CommentDto> Replies { get; set; }

        public virtual AttachmentDto Attachment { get; set; }

        public DateTime? CreateDate { get; set; }

        public DateTime? UpdateDate { get; set; }

        public AccountDto Creator { get; set; }

        public AccountDto Updater { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblCmComment, CommentDto>().ReverseMap();
        }
    }

    public class CommentCreateDto : IMapFrom, IDto
    {
        public int? PId { get; set; }

        public string Type { get; set; }

        public string Content { get; set; }

        public Guid? AttachmentId { get; set; }

        public Guid ReferenceId { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblCmComment, CommentCreateDto>().ReverseMap();
        }
    }

    public class CommentUpdateDto : IMapFrom, IDto
    {
        public int Id { get; set; }

        public string Type { get; set; }

        public string Content { get; set; }

        public Guid? AttachmentId { get; set; }

        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblCmComment, CommentUpdateDto>().ReverseMap();
        }
    }
}
