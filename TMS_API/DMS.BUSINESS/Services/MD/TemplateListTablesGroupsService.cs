using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.MD;
using DMS.CORE;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.MD
{
    public interface ITemplateListTablesGroupsService : IGenericService<TblMdTemplateListTablesGroups,TemplateListTablesGroupsDto>
    {
        Task<IList<TemplateListTablesGroupsDto>> GetAll(BaseMdFilter filter);
    }
    public class TemplateListTablesGroupsService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdTemplateListTablesGroups, TemplateListTablesGroupsDto>(dbContext, mapper), ITemplateListTablesGroupsService
    {
        public async Task<IList<TemplateListTablesGroupsDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.tblMdTemplateListTablesGroups.AsQueryable();
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
                var query = _dbContext.tblMdTemplateListTablesGroups.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Code.ToString().Contains(filter.KeyWord) || x.Id.ToString().Contains(filter.KeyWord) ||
                                            x.Name.ToString().Contains(filter.KeyWord));
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
