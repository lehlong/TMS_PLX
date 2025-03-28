using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Filter.BU;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using Common;
using Microsoft.EntityFrameworkCore;

namespace DMS.BUSINESS.Services.CM.Comment
{
    public interface ICommentService : IGenericService<TblCmComment, CommentDto>
    {
        Task<PagedResponseDto> GetCommentsByReference(CommentFilter filter);
        Task<CommentDto> Add(CommentCreateDto dto);
        Task Update(CommentUpdateDto dto);
        Task Delete(int Id);
    }

    public class CommentService(AppDbContext dbContext, IMapper mapper) : GenericService<TblCmComment, CommentDto>(dbContext, mapper), ICommentService
    {
        public async Task<PagedResponseDto> GetCommentsByReference(CommentFilter filter)
        {
            try
            {
                var objComments = filter.ReferenceId != null ?
                    _dbContext.TblCmModuleComment.Where(x => x.ReferenceId == filter.ReferenceId).ToList() :
                    [.. _dbContext.TblCmModuleComment];

                var commentIds = objComments.Select(x => x.CommentId);

                var comments = _dbContext.TblCmComment
                                .Include(x => x.Replies)
                                    .ThenInclude(x => x.Attachment)
                                .Include(x => x.Replies)
                                    .ThenInclude(x => x.Creator)
                                .Include(x => x.Attachment)
                                .Include(x => x.Creator)
                                .Where(x => commentIds.Contains(x.Id))
                                .Where(x => x.PId == null)
                                .OrderByDescending(x => x.CreateDate);

                return await Paging(comments, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<CommentDto> Add(CommentCreateDto dto)
        {
            try
            {
                await _dbContext.Database.BeginTransactionAsync();

                var entity = _mapper.Map<TblCmComment>(dto);

                entity.Replies = [];

                var entityResult = await _dbContext.TblCmComment.AddAsync(entity);
                await _dbContext.SaveChangesAsync();

                var newCommentModule = new TblCmModuleComment
                {
                    ReferenceId = dto.ReferenceId,
                    CommentId = entity.Id
                };

                await _dbContext.TblCmModuleComment.AddAsync(newCommentModule);
                await _dbContext.SaveChangesAsync();

                var dtoResult = _mapper.Map<CommentDto>(entityResult.Entity);
                await _dbContext.Database.CommitTransactionAsync();

                return dtoResult;
            }
            catch (Exception ex)
            {
                await _dbContext.Database.RollbackTransactionAsync();
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task Update(CommentUpdateDto dto)
        {
            try
            {
                await _dbContext.Database.BeginTransactionAsync();

                var entityInDB = _dbContext.TblCmComment
                                .Include(x => x.Replies)
                                .Include(x => x.Attachment)
                                .Include(x => x.Creator)
                                .FirstOrDefault(x => x.Id == dto.Id);

                if (entityInDB == null)
                {
                    Status = false;
                    MessageObject.Code = "0003";
                    return;
                }

                _mapper.Map(dto, entityInDB);

                _dbContext.TblCmComment.Update(entityInDB);
                await _dbContext.SaveChangesAsync();

                await _dbContext.Database.CommitTransactionAsync();
            }
            catch (Exception ex)
            {
                await _dbContext.Database.RollbackTransactionAsync();
                Status = false;
                Exception = ex;
                return;
            }
        }

        public async Task Delete(int Id)
        {
            try
            {
                var commentToDelete = _dbContext.TblCmComment
                                    .Include(x => x.Replies)
                                    .Include(x => x.Attachment)
                                    .FirstOrDefault(x => x.Id == Id);

                _dbContext.RemoveRange(commentToDelete);
                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return;
            }
        }
    }
}
