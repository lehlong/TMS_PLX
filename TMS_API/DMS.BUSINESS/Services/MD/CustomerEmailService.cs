using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Filter.MD;
using DMS.CORE;
using DMS.CORE.Entities.MD;
using Microsoft.EntityFrameworkCore;

namespace DMS.BUSINESS.Services.MD
{
    public interface ICustomerEmailService : IGenericService<TblMdCustomerEmail, CustomerEmailDto>
    {
        Task<IList<CustomerEmailDto>> GetAll(BaseMdFilter filter);
        Task<byte[]> Export(BaseMdFilter filter);

    }
    public class CustomerEmailService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdCustomerEmail, CustomerEmailDto>(dbContext, mapper), ICustomerEmailService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCustomerEmail.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Email.Contains(filter.KeyWord) || 
                    x.Code.ToString().Contains(filter.KeyWord) ||
                    x.CustomerCode == (filter.KeyWord)
                    );
                }
                //if (filter.IsActive.HasValue)
                //{
                //    query = query.Where(x => x.IsActive == filter.IsActive);
                //}
                return await Paging(query, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<IList<CustomerEmailDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCustomerEmail.AsQueryable();
                //if (filter.IsActive.HasValue)
                //{
                //    query = query.Where(x => x.IsActive == filter.IsActive);
                //}
                return await base.GetAllMd(query, filter);

            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task<byte[]> Export(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCustomerEmail.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Email.Contains(filter.KeyWord));
                }
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                var data = await base.GetAllMd(query, filter);
                int i = 1;
                data.ForEach(x =>
                {
                    x.OrdinalNumber = i++;
                });
                return await ExportExtension.ExportToExcel(data);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
    }
}
