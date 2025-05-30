﻿using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.AD
{
    public interface IConfigMailService : IGenericService<TblEmailConfig, EmailConfigDto>
    {
        Task<IList<EmailConfigDto>> GetAll(BaseMdFilter filter);
        //Task<byte[]> Export(BaseMdFilter filter);
    }
    public class ConfigMailService(AppDbContext dbContext, IMapper mapper) : GenericService<TblEmailConfig, EmailConfigDto>(dbContext, mapper), IConfigMailService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblEmailConfig.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Email.Contains(filter.KeyWord));
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
        public async Task<IList<EmailConfigDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblEmailConfig.AsQueryable();
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
        //public async Task<byte[]> Export(BaseMdFilter filter)
        //{
        //    try
        //    {
        //        var query = _dbContext.TblEmailConfig.AsQueryable();
        //        if (!string.IsNullOrWhiteSpace(filter.KeyWord))
        //        {
        //            query = query.Where(x => x.Email.Contains(filter.KeyWord));
        //        }
        //        if (filter.IsActive.HasValue)
        //        {
        //            query = query.Where(x => x.IsActive == filter.IsActive);
        //        }
        //        var data = await base.GetAllMd(query, filter);
        //        int i = 1;
        //        data.ForEach(x =>
        //        {
        //            x.OrdinalNumber = i++;
        //        });
        //        return await ExportExtension.ExportToExcel(data);
        //    }
        //    catch (Exception ex)
        //    {
        //        Status = false;
        //        Exception = ex;
        //        return null;
        //    }
        //}

    }
}
