using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Entities.MD;

namespace DMS.BUSINESS.Services.MD
{
    public interface ICuocVanChuyenListService : IGenericService<TblMdCuocVanChuyenList, CuocVanChuyenListDto>
    {
        Task<IList<CuocVanChuyenListDto>> GetAll(BaseMdFilter filter);
    }
    class CuocVanChuyenListService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdCuocVanChuyenList, CuocVanChuyenListDto>(dbContext, mapper), ICuocVanChuyenListService
    {
        public async Task<IList<CuocVanChuyenListDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCuocVanChuyenList.AsQueryable();
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

        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCuocVanChuyenList.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Name.Contains(filter.KeyWord));
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
    }
}


