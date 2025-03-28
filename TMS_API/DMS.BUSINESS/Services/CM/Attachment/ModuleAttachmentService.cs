using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Common.Enum;
using DMS.BUSINESS.Common.Util;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.BU.Attachment;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;

namespace DMS.BUSINESS.Services.BU
{
    public interface IModuleAttachmentService : IGenericService<TblCmModuleAttachment, ModuleAttachmentDto>
    {
        Task<object> Upload(IFormFile file, string moduleType, Guid? referenceId, Guid? deleteAttachmentId = null);
        Task<List<ModuleAttachmentDto>> GetByReferenceId(Guid refId);
        Task<object> UploadList(List<IFormFile> files, string? moduleType, Guid? referenceId, List<Guid>? deleteAttachmentId = null);
    }
    public class ModuleAttachmentService(AppDbContext dbContext, IMapper mapper, IAttachmentManagerService attachmentManager) : GenericService<TblCmModuleAttachment, ModuleAttachmentDto>(dbContext, mapper), IModuleAttachmentService
    {
        private readonly IAttachmentManagerService _attachmentManager = attachmentManager;

        public override async Task<ModuleAttachmentDto> GetById(object id)
        {
            try
            {
                var data = await _dbContext.TblBuModuleAttachment.Include(x => x.Attachment)
                    .FirstOrDefaultAsync(x => x.Id == (Guid)id);

                return _mapper.Map<ModuleAttachmentDto>(data);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<List<ModuleAttachmentDto>> GetByReferenceId(Guid refId)
        {
            try
            {
                var data = await _dbContext.TblBuModuleAttachment.Include(x => x.Attachment)
                    .Where(x => x.ReferenceId == refId).OrderByDescending(x => x.CreateDate).ToListAsync();

                return _mapper.Map<List<ModuleAttachmentDto>>(data);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<object> Upload(IFormFile file, string? moduleType, Guid? referenceId, Guid? deleteAttachmentId)
        {
            if (deleteAttachmentId != null)
            {
                await _attachmentManager.Delete([deleteAttachmentId.Value]);
            }

            if (!Enum.TryParse(moduleType, out ModuleType mdType))
            {
                this.Status = false;
                this.MessageObject.Code = "2005";
                mdType = ModuleType.PARTNER;
                //return null;
            }

            if (file.Length < 0)
            {
                this.Status = false;
                this.MessageObject.Code = "3001";
                return null;
            }

            using var ms = new MemoryStream();
            file.CopyTo(ms);
            var fileBytes = ms.ToArray();

            var uploadResult = await _attachmentManager.UploadModuleAttachment(fileBytes, file.FileName, Path.GetExtension(file.FileName), FileUtil.GetFileType(Path.GetExtension(file.FileName)), mdType, referenceId);

            this.Status = uploadResult.Status;

            if (uploadResult.Exception != null)
            {
                this.Exception = uploadResult.Exception;
            }

            if (uploadResult.MessageObject != null)
            {
                this.MessageObject = uploadResult.MessageObject;

            }

            return uploadResult.Data;
        }

        public async Task<object> UploadList(List<IFormFile> files, string? moduleType, Guid? referenceId, List<Guid>? deleteAttachmentIds)
        {
            if(deleteAttachmentIds != null && deleteAttachmentIds.Count != 0)
            {
                await _attachmentManager.Delete(deleteAttachmentIds);
            }

            if (!Enum.TryParse(moduleType, out ModuleType mdType))
            {
                this.Status = false;
                this.MessageObject.Code = "2005";
                mdType = ModuleType.PARTNER;
                //return null;
            }

            List<BatchUploadDto> datas = [];

            foreach (var file in files)
            {
                using var ms = new MemoryStream();
                file.CopyTo(ms);
                var fileBytes = ms.ToArray();
                datas.Add(new BatchUploadDto()
                {
                    ByteData = fileBytes,
                    FileExtension = Path.GetExtension(file.FileName),
                    FileName = file.FileName,
                    FileType = FileUtil.GetFileType(Path.GetExtension(file.FileName))
                });
            }

            var uploadResult = await _attachmentManager.BatchUploadModuleAttachment(datas, mdType, referenceId);

            this.Status = uploadResult.Status;

            if (uploadResult.Exception != null)
            {
                this.Exception = uploadResult.Exception;
            }

            if (uploadResult.MessageObject != null)
            {
                this.MessageObject = uploadResult.MessageObject;

            }

            return uploadResult.Data;
        }
    }
}
