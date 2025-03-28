using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.IN;
using ICSharpCode.SharpZipLib.Zip;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.BU
{
    public interface IDiscountInformationListService : IGenericService<TblBuDiscountInformationList, DiscountInformationListDto>
    {
        Task<IList<DiscountInformationListDto>> GetAll(BaseMdFilter filter);
        Task InsertData(CompetitorModel model);
        Task<CompetitorModel> BuildObjectCreate(string code);
        Task<CompetitorModel> GetDataByCode(string code);
    }

    public class DiscountInformationListService(AppDbContext dbContext, IMapper mapper) : GenericService<TblBuDiscountInformationList, DiscountInformationListDto>(dbContext, mapper), IDiscountInformationListService
    {
        public async Task<IList<DiscountInformationListDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblBuDiscountInformationList.AsQueryable();
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

        public async Task InsertData(CompetitorModel model)
        {
            try 
            {
                _dbContext.TblBuDiscountInformationList.Add(model.Header);
                foreach (var g in model.goodss)
                {
                    _dbContext.TblInDiscountCompetitor.AddRange(g.HS);
                    _dbContext.TblInDiscountCompany.AddRange(g.DiscountCompany);
                }
                ;
                await _dbContext.SaveChangesAsync();
            }
            catch(Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }

        public async Task<CompetitorModel> BuildObjectCreate(string code)
        {

            var dateTimeNow = DateTime.Now;
            //var fdate = await _dbContext.TblBuDiscountInformationList.Where(x => x.Code == code).Select(x => x.FDate).FirstOrDefaultAsync();
            try
            {

                var obj = new CompetitorModel();
                obj.HeaderName = await _dbContext.TblBuCalculateResultList.Where(x => x.Code == code).Select(x => x.Name).FirstOrDefaultAsync() ?? "";
                obj.Header.Code = code;
                obj.Header.Name = await _dbContext.TblBuDiscountInformationList.Where(x => x.Code == code).Select(x => x.Name).FirstOrDefaultAsync() ?? "";
                //obj.Header.FDate = DateTime.Now;
                obj.Header.FDate = await _dbContext.TblBuCalculateResultList
                .Where(x => x.Code == code)
                .Select(x => (DateTime?)x.FDate)
                .FirstOrDefaultAsync() ?? dateTimeNow;

                obj.Header.IsActive = true;

                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.Code).ToListAsync();
                var lstCompetitor = await _dbContext.TblMdCompetitor.OrderBy(x => x.Code).ToListAsync();
                var lstDiscountInformation = await _dbContext.TblInDiscountCompetitor.Where(x => x.HeaderCode == code).ToArrayAsync();

                foreach (var g in lstGoods)
                {
                    var goods = new GOODSs();
                    goods.Code = g.Code;
                    foreach (var c in lstCompetitor)
                    {
                        goods.HS.Add(new TblInDiscountCompetitor
                        {
                            Code = lstDiscountInformation.Where(x => x.GoodsCode == g.Code && x.CompetitorCode == c.Code).Select(x => x.Code).FirstOrDefault() ?? Guid.NewGuid().ToString(),
                            HeaderCode = obj.Header.Code,
                            GoodsCode = g.Code,
                            Discount = lstDiscountInformation.Where(x => x.GoodsCode == g.Code && x.CompetitorCode == c.Code).Sum(x => x.Discount ?? 0.00M),
                            CompetitorCode = c.Code,
                            IsActive = true,
                        });               
                    }

                    var discount = await _dbContext.TblInDiscountCompany
                        .Where(x => x.HeaderCode == code && x.GoodsCode == g.Code)
                        .Select(x => x.Discount)
                        .FirstOrDefaultAsync();

                    if (discount == null)
                    {
                        var codeDiscountInformationList = await _dbContext.TblBuDiscountInformationList.OrderByDescending(x => x.FDate).Select(x => x.Code).FirstOrDefaultAsync();

                        discount = await _dbContext.TblInDiscountCompany
                         .Where(x => x.GoodsCode == g.Code && x.HeaderCode == codeDiscountInformationList)
                         .Select(x => x.Discount)
                         .FirstOrDefaultAsync();
                        Console.WriteLine($"Tìm thấy Discount: {discount}");
                        Console.WriteLine($"Tìm thấy Code: {codeDiscountInformationList}");
                    }

                    

                    goods.DiscountCompany.Add(new TblInDiscountCompany
                    {
                        Code = await _dbContext.TblInDiscountCompany.Where(x => x.HeaderCode == code && x.GoodsCode == g.Code).Select(x => x.Code).FirstOrDefaultAsync() ?? Guid.NewGuid().ToString(),
                        HeaderCode = obj.Header.Code,
                        Discount = discount ?? 0,
                        GoodsCode = g.Code,
                    });
                    obj.goodss.Add(goods);
                }
                return obj;
            }
            catch
            {
                return new CompetitorModel();
            }
        }
        public async Task<CompetitorModel> GetDataByCode(String Code)
        {

            return null;
        }
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblBuDiscountInformationList.AsQueryable();

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

    public class CompetitorModel
    {
        public string? HeaderName { get; set; }

        public TblBuDiscountInformationList Header { get; set; } = new TblBuDiscountInformationList();
        public List<GOODSs> goodss { get; set; } = new List<GOODSs>();
    }
    public class GOODSs
    {
        public string? Code { get; set; }
        public List<TblInDiscountCompetitor> HS { get; set; } = new List<TblInDiscountCompetitor>();

        public List<TblInDiscountCompany> DiscountCompany { get; set; } = new List<TblInDiscountCompany>();
    }
}
