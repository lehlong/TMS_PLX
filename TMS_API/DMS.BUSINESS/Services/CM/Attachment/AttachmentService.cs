using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace DMS.BUSINESS.Services.BU.Attachment
{
    public interface IAttachmentService : IGenericService<TblCmAttachment, AttachmentDto>
    {
        Task<(byte[], string)> Download(Guid attachmentId);
    }

    public class AttachmentService(AppDbContext dbContext, IMapper mapper, IConfiguration configuration) : GenericService<TblCmAttachment, AttachmentDto>(dbContext, mapper), IAttachmentService
    {
        public async Task<(byte[], string)> Download(Guid attachmentId)
        {
            var attachment = await _dbContext.TblBuAttachment.FirstOrDefaultAsync(x => x.Id == attachmentId);

            if (attachment == null)
            {
                this.Status = false;
                this.MessageObject.Code = "0003";
            }

            var uploadUrl = configuration.GetSection("Path:Upload").Value;

            var path = Path.Combine(uploadUrl, attachment.Url.Replace("\\","/"));

            return ((File.ReadAllBytes(path), Path.GetFileName(path)));
        }
    }
}
