using Common;
using DMS.BUSINESS.Common.Enum;
using DMS.BUSINESS.Dtos.BU;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Storage;
using Microsoft.Extensions.Configuration;

namespace DMS.BUSINESS.Services.BU.Attachment
{
    public interface IAttachmentManagerService
    {
        Task<ServiceResponseDto> UploadModuleAttachment(byte[] data, string fileName, string extension, string fileType, ModuleType? moduleType, Guid? RefId = null);
        Task<ServiceResponseDto> BatchUploadModuleAttachment(List<BatchUploadDto> datas, ModuleType moduleType, Guid? RefId = null);
        Task<ServiceResponseDto> Delete(List<Guid> attachmentIds);
    }
    public class AttachmentManagerService : IAttachmentManagerService
    {
        private readonly AppDbContext _dbContext;
        private readonly IConfiguration _configuration;
        private readonly string _baseUploadUrl;
        public AttachmentManagerService(AppDbContext context, IConfiguration configuration)
        {
            _dbContext = context;
            _configuration = configuration;
            _baseUploadUrl = _configuration.GetSection("Path:Uploads").Value ?? string.Empty;
        }

        public async Task<ServiceResponseDto> UploadModuleAttachment(byte[] data, string fileName, string extension, string fileType, ModuleType? moduleType, Guid? RefId = null)
        {
            if (RefId == null) RefId = Guid.NewGuid();

            string physicalFileName = Guid.NewGuid().ToString();

            //var uploadPath = Path.Combine(moduleType.ToString(), DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString(), $"{physicalFileName}{extension}");
            var rootUploadPath = Path.Combine("Uploads");
            var uploadPath = Path.Combine(
                rootUploadPath,
                moduleType.ToString(),
                DateTime.Now.Year.ToString(),
                DateTime.Now.Month.ToString(),
                DateTime.Now.Day.ToString(),
                $"{physicalFileName}{extension}"
            );

            // Tạo thư mục nếu chưa tồn tại
            Directory.CreateDirectory(Path.GetDirectoryName(uploadPath));
            var uploadResult = await Upload(data, uploadPath);

            if (uploadResult.Status == false) return uploadResult;

            IDbContextTransaction transaction = null;

            if (_dbContext.Database.CurrentTransaction == null)
            {
                transaction = _dbContext.Database.BeginTransactionAsync().Result;
            }

            try
            {
                var attachment = new TblCmAttachment
                {
                    Id = Guid.NewGuid(),
                    Url = uploadPath,
                    Extension = extension,
                    Size = data.Length / 1024.0,
                    Type = fileType,
                    Name = fileName,
                };

                await _dbContext.TblBuAttachment.AddAsync(attachment);
                await _dbContext.SaveChangesAsync();

                var moduleAttachment = new TblCmModuleAttachment()
                {
                    AttachmentId = attachment.Id,
                    ReferenceId = RefId,
                    ModuleType = moduleType.ToString(),
                };

                await _dbContext.TblBuModuleAttachment.AddAsync(moduleAttachment);
                await _dbContext.SaveChangesAsync();

                transaction?.CommitAsync();
                return new()
                {
                    Data = new
                    {
                        AttachmentId = attachment.Id,
                        Url = (string)uploadResult.Data,
                        ReferenceId = RefId
                    },
                    Status = true,
                };
            }
            catch (Exception ex)
            {
                transaction?.RollbackAsync();
                return new()
                {
                    Data = null,
                    Status = false,
                    Exception = ex,
                };
            }
        }

        public async Task<ServiceResponseDto> BatchUploadModuleAttachment(List<BatchUploadDto> datas, ModuleType moduleType, Guid? RefId = null)
        {
            if (datas == null || datas.Count == 0)
            {
                return new()
                {
                    Data = null,
                    Status = true
                };
            }

            var uploadPathInDay = Path.Combine(moduleType.ToString(), DateTime.Now.Year.ToString(), DateTime.Now.Month.ToString(), DateTime.Now.Day.ToString());


            if (RefId == null) RefId = Guid.NewGuid();

            IDbContextTransaction transaction = null;
            if (_dbContext.Database.CurrentTransaction == null)
            {
                transaction = _dbContext.Database.BeginTransactionAsync().Result;
            }
            try
            {
                List<TblCmAttachment> attachments = [];
                List<TblCmModuleAttachment> moduleAttachments = [];
                foreach (var data in datas)
                {
                    string physicalFileName = Guid.NewGuid().ToString();

                    var uploadPath = Path.Combine(uploadPathInDay, $"{physicalFileName}{data.FileExtension}");

                    var uploadResult = await Upload(data.ByteData, uploadPath);
                    if (!uploadResult.Status) continue;

                    var attachment = new TblCmAttachment
                    {
                        Url = uploadPath,
                        Extension = data.FileExtension,
                        Size = data.ByteData.Length / 1024.0,
                        Type = data.FileType.ToString(),
                        Name = data.FileName,
                    };

                    attachments.Add(attachment);
                }

                await _dbContext.TblBuAttachment.AddRangeAsync(attachments);
                await _dbContext.SaveChangesAsync();

                foreach (var attachment in attachments)
                {
                    var moduleAttachment = new TblCmModuleAttachment()
                    {
                        AttachmentId = attachment.Id,
                        ReferenceId = RefId,
                        ModuleType = moduleType.ToString(),
                    };

                    moduleAttachments.Add(moduleAttachment);
                }

                await _dbContext.TblBuModuleAttachment.AddRangeAsync(moduleAttachments);
                await _dbContext.SaveChangesAsync();

                transaction?.CommitAsync();
                return new()
                {
                    Data = new
                    {
                        Attachment = attachments.Select(x => new
                        {
                            x.Id,
                        }).ToList(),
                        ReferenceId = RefId
                    },
                    Status = true,
                };

            }
            catch (Exception ex)
            {
                transaction?.RollbackAsync();
                return new()
                {
                    Data = null,
                    Status = false,
                    Exception = ex,
                };
            }
        }

        private async Task<ServiceResponseDto> Upload(byte[] data, string path)
        {
            try
            {
                if (data == null || data.Length == 0)
                {
                    return new()
                    {
                        Data = null,
                        MessageObject = new MessageObject() { Code = "3000" },
                        Status = false
                    };
                }

                var fullPath = Path.Combine(_baseUploadUrl, path);

                string directoryPath = Path.GetDirectoryName(fullPath);

                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }

                await File.WriteAllBytesAsync(fullPath, data);

                return new()
                {
                    Status = true,
                };
            }
            catch (Exception ex)
            {
                return new()
                {
                    Status = false,
                    Exception = ex,
                };
            }
        }

        public async Task<ServiceResponseDto> Delete(List<Guid> attachmentIds)
        {
            try
            {
                var moduleAttachments = await _dbContext.TblBuModuleAttachment.Where(x => attachmentIds.Contains(x.AttachmentId)).ToListAsync();
                _dbContext.TblBuModuleAttachment.RemoveRange(moduleAttachments);
                await _dbContext.SaveChangesAsync();
                var attachments = await _dbContext.TblBuAttachment.Where(x => attachmentIds.Contains(x.Id)).ToListAsync();
                var uploadUrl = _configuration.GetSection("Path:Upload").Value;
                foreach (var attachment in attachments)
                {
                    var path = Path.Combine(uploadUrl, attachment.Url.Replace("\\", "/"));
                    File.Delete(path);
                }
                _dbContext.TblBuAttachment.RemoveRange(attachments);
                await _dbContext.SaveChangesAsync();
                return new()
                {
                    Status = true,
                };
            }
            catch (Exception ex)
            {
                return new()
                {
                    Status = false,
                    Exception = ex,
                };
            }
        }
    }
}
