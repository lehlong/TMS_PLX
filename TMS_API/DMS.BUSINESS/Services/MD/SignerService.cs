using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.BU;
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
    public interface ISignerService : IGenericService<TblMdSigner, SignerDto>
    {
        Task<IList<SignerDto>> GetAll(BaseMdFilter filter);
        Task UpdateSigner(TblMdSigner data);
        Task<TblMdSigner> Insert(TblMdSigner signer);
    }
    class SignerService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdSigner, SignerDto>(dbContext, mapper), ISignerService
    {
        public async Task<IList<SignerDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdSigner.AsQueryable();
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
                var query = _dbContext.TblMdSigner.AsQueryable();

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

        public async Task<TblMdSigner> Insert(TblMdSigner signer)
        {
            try
            {
                if (signer.IsSelect == true)
                {
                    var lstSigner = await _dbContext.TblMdSigner.Where(x => x.Type == signer.Type).ToListAsync();

                    foreach (var item in lstSigner)
                    {
                        item.IsSelect = false;
                    }
                }
                _dbContext.TblMdSigner.Add(signer);
                await _dbContext.SaveChangesAsync();

                return signer;
            }
            catch (Exception ex)
            {
                return new TblMdSigner();
            }
        }

        public async Task UpdateSigner(TblMdSigner data)
        {
            try
            {
                var currentSelectedSigner = await _dbContext.TblMdSigner
                    .FirstOrDefaultAsync(x => x.Type == data.Type && x.IsSelect == true);

                if (currentSelectedSigner != null && currentSelectedSigner.Code != data.Code)
                {
                    currentSelectedSigner.IsSelect = false;
                }

                data.IsSelect = true;

                _dbContext.TblMdSigner.Update(data);

                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                //return null;
            }
        }
    }
}
