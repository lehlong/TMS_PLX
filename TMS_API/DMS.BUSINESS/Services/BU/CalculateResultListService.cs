using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.IN;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.BU
{
    public interface ICalculateResultListService : IGenericService<TblBuCalculateResultList, CalculateResultListDto>
    {
        Task<IList<CalculateResultListDto>> GetAll(BaseMdFilter filter);
        Task<byte[]> Export(BaseMdFilter filter);
        Task InsertData(InsertModel model);
        Task<InsertModel> GetObjectCreate();
    }
    public class CalculateResultListService(AppDbContext dbContext, IMapper mapper) : GenericService<TblBuCalculateResultList, CalculateResultListDto>(dbContext, mapper), ICalculateResultListService
    {
        public async Task InsertData(InsertModel model)
        {
            try
            {
                _dbContext.TblBuCalculateResultList.Add(model.Header);
                _dbContext.TblInHeSoMatHang.AddRange(model.HS1);
                _dbContext.TblInVinhCuaLo.AddRange(model.HS2);
                _dbContext.TbLInCustomerFob.AddRange(model.FOB);
                var h = new TblBuHistoryAction()
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = model.Header.Code,
                    Action = "Tạo mới",
                };
                _dbContext.TblBuHistoryAction.Add(h);
                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblBuCalculateResultList.AsQueryable();

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
        public async Task<IList<CalculateResultListDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblBuCalculateResultList.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                if (filter.KeyWord != null)
                {
                    query = query.Where(x => x.Status == filter.KeyWord);

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
        public async Task<byte[]> Export(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblBuCalculateResultList.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Name.Contains(filter.KeyWord));
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

        public async Task<InsertModel> GetObjectCreate()
        {
            try
            {
                var obj = new InsertModel();
                obj.Header.Code = Guid.NewGuid().ToString();
                obj.Header.SignerCode = "TongGiamDoc";
                obj.Header.QuyetDinhSo = "";
                obj.Header.IsActive = true;
                obj.Header.FDate = DateTime.Now;
                obj.Header.CreateDate = DateTime.Now;
                obj.Header.UpdateDate = DateTime.Now;
                obj.Header.IsDeleted = false;
                obj.Header.DeleteDate = DateTime.Now;
                obj.Header.Name = "";
                obj.Header.Status = "01";

                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive==true).OrderBy(x => x.CreateDate).ToListAsync();
                foreach (var g in lstGoods)
                {
                    obj.HS1.Add(new TblInHeSoMatHang
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = obj.Header.Code,
                        GoodsCode = g.Code,
                        ThueBvmt = g.ThueBvmt,
                        HeSoVcf = g.Vfc,
                        IsActive = true,
                        FromDate = DateTime.Now,
                        ToDate = DateTime.Now,
                        CreateDate = DateTime.Now,
                        UpdateDate = DateTime.Now
                    });
                    obj.HS2.Add(new TblInVinhCuaLo
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = obj.Header.Code,
                        GoodsCode = g.Code,
                        MtsV1 = g.MtsV1,
                        IsActive = true,
                        FromDate = DateTime.Now,
                        ToDate = DateTime.Now,
                        CreateDate = DateTime.Now,
                        UpdateDate = DateTime.Now
                    });
                }
                var lstCustomerSpecial = await _dbContext.TblMdCustomer.Where(x => x.IsActive == true && x.Fob > 0).ToListAsync();
                foreach (var c in lstCustomerSpecial)
                {
                    obj.FOB.Add(new TbLInCustomerFob
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = obj.Header.Code,
                        CustomerName = c.Name,
                        CustomerCode = c.Code,
                        Fob = c.Fob,
                        IsActive = true,
                        CreateDate = DateTime.Now,
                    });
                }
                return obj;
            }
            catch
            {
                return new InsertModel();
            }
        }

    }

    public class InsertModel
    {
        public TblBuCalculateResultList Header { get; set; } = new TblBuCalculateResultList();
        public List<TblInHeSoMatHang> HS1 { get; set; } = new List<TblInHeSoMatHang>();
        public List<TblInVinhCuaLo> HS2 { get; set; } = new List<TblInVinhCuaLo>();

        public List<TbLInCustomerFob> FOB { get; set; } = new List<TbLInCustomerFob>();
        public Status Status { get; set; } = new Status();
    }
    public class Status
    {
        public string? Code { get; set; }
        public string? Contents { get; set; }
    }
}
