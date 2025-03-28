using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Entities.MD;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.MD
{
    public interface IPeriodTimeService : IGenericService<TblMdPeriodTime, PeriodTimeDto>
    {
        Task<IList<PeriodTimeDto>> GetAll(BaseMdFilter filter);
        Task<bool> ChangeDefaultStatus(int timeYear);
        Task<bool> ChangeIsClosedStatus(int timeYear);
    }
    public class PeriodTimeService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdPeriodTime, PeriodTimeDto>(dbContext, mapper), IPeriodTimeService
    {
        public override async Task<PagedResponseDto> Search (BaseFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdPeriodTime.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.TimeYear.ToString().Contains(filter.KeyWord));
                }
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                return await Paging(query, filter);

            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task<IList<PeriodTimeDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdPeriodTime.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                return await base.GetAllMd(query, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public override Task<PeriodTimeDto> Add(IDto dto)
        {
            var model = dto as PeriodTimeDto;
            model.IsDefault = false;
            model.IsClosed = false;
            return base.Add(dto);
        }
        public async Task<bool> ChangeDefaultStatus(int timeYear)
        {
            try
            {
                using var transaction = await _dbContext.Database.BeginTransactionAsync();
                // Tìm bản ghi cần thay đổi
                var periodTime = await _dbContext.tblMdPeriodTime.FindAsync(timeYear);
                if (periodTime == null)
                {
                    return false; // Không tìm thấy bản ghi
                }
                // Nếu bản ghi đã là mặc định, không cần thay đổi
                if (periodTime.IsDefault)
                {
                    return true;
                }
                // Đặt tất cả các bản ghi khác thành false
                await _dbContext.tblMdPeriodTime.Where(p => p.IsDefault)
                    .ForEachAsync(p => p.IsDefault = false);
                // Đặt bản ghi hiện tại thành true
                periodTime.IsDefault = true;
                await _dbContext.SaveChangesAsync();
                await transaction.CommitAsync();

                return true;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return false;
            }
        }
        public async Task<bool> ChangeIsClosedStatus(int timeYear)
        {
            try
            {
                var periodTime = await _dbContext.tblMdPeriodTime.FindAsync(timeYear);
                if (periodTime == null)
                {
                    return false;
                }

                // Đảo ngược trạng thái IsClosed
                periodTime.IsClosed = !periodTime.IsClosed;

                await _dbContext.SaveChangesAsync();

                return true;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return false;
            }
        }
    }
}
