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
    public interface IGiaGiaoTapDoanService : IGenericService<TblMdGiaGiaoTapDoan, GiaGiaoTapDoanDto>
    {
        Task<GgtdModel> GetAll();
        Task<byte[]> Export(BaseMdFilter filter);
        Task<GgtdModel> BuildDataCreate();
        Task<GgtdModel> InsertData(GgtdModel model);
        Task<GgtdModel> UpdateData(GgtdModel model);
    }
    public class GiaGiaoTapDoanService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdGiaGiaoTapDoan, GiaGiaoTapDoanDto>(dbContext, mapper), IGiaGiaoTapDoanService
    {
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdGiaGiaoTapDoan.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Code.Contains(filter.KeyWord));
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
        public async Task<GgtdModel> GetAll()
        {
            try
            {
                var lstGgtdL = _dbContext.TblMdGiaGiaoTapDoanList.Where(x => x.IsActive == null).OrderByDescending(x => x.FDate).ToList();
                var lstGgtd = await _dbContext.TblMdGiaGiaoTapDoan.ToListAsync();

                var data = new GgtdModel();

                foreach (var item in lstGgtdL)
                {
                    var ggtdlModel = new GgtdModel();
                    //List<TblMdGiaGiaoTapDoan> listGgtd = ;
                    ggtdlModel.GgtdlHeader = item;
                    ggtdlModel.Ggtd = lstGgtd.Where(x => x?.GgtdlCode == item.Code).ToList();
                    
                    
                    data = ggtdlModel;
                }

                return data;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<GgtdModel> BuildDataCreate()
        {
            try
            {
                var OldGgtdList = await _dbContext.TblMdGiaGiaoTapDoanList
                                                .OrderByDescending(x => x.FDate) 
                                                .FirstOrDefaultAsync();
                List<TblMdGiaGiaoTapDoan> LstOldGgtd;
                if (OldGgtdList == null || OldGgtdList.Code == null)
                {
                    LstOldGgtd = new List<TblMdGiaGiaoTapDoan>();
                }
                else
                {
                    LstOldGgtd = await _dbContext.TblMdGiaGiaoTapDoan
                        .Where(x => x.GgtdlCode == OldGgtdList.Code)
                        .ToListAsync();
                }
                //var LstOldGgtd = await _dbContext.TblMdGiaGiaoTapDoan.Where(x => x.GgtdlCode == OldGgtdList.Code).ToListAsync();
                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.Code).ToListAsync();

                var ggtdModel = new GgtdModel();


                ggtdModel.oldHeaderGgtd = OldGgtdList == null ? "" : OldGgtdList.Code;

                ggtdModel.GgtdlHeader.Code = Guid.NewGuid().ToString();
                ggtdModel.GgtdlHeader.FDate = DateTime.Now;
                ggtdModel.GgtdlHeader.Name = "";
                foreach (var g in lstGoods)
                {
                    var ggtd = new TblMdGiaGiaoTapDoan();
                    ggtd.Code = Guid.NewGuid().ToString();
                    ggtd.GoodsCode = g.Code;
                    ggtd.GgtdlCode = ggtdModel.GgtdlHeader.Code;
                    ggtd.NewPrice = 0;
                    ggtd.OldPrice = LstOldGgtd.Where(x => x.GgtdlCode == OldGgtdList.Code && x.GoodsCode == g.Code).Select(x => x.NewPrice).SingleOrDefault() ?? 0;
                    
                    ggtdModel.Ggtd.Add(ggtd);
                }

                return ggtdModel;
            }
            catch
            {
                return new GgtdModel();
            }
        }

        public async Task<GgtdModel> InsertData(GgtdModel model)
        {
            try
            {
                var exists = await _dbContext.TblMdGiaGiaoTapDoanList
                    .AnyAsync(item => item.FDate > model.GgtdlHeader.FDate);
                var OldGgtdList = await _dbContext.TblMdGiaGiaoTapDoanList.Where(x => x.Code == model.oldHeaderGgtd).FirstOrDefaultAsync();

                OldGgtdList.IsActive = false;
                if (exists)
                {
                    model.oldHeaderGgtd = "false";
                    return model;
                }
                else
                {
                    // Không có giá trị nào thỏa mãn
                    _dbContext.TblMdGiaGiaoTapDoanList.Add(model.GgtdlHeader);
                    _dbContext.TblMdGiaGiaoTapDoan.AddRange(model.Ggtd);
                
                    await _dbContext.SaveChangesAsync();
                    return model;
                }
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task<GgtdModel> UpdateData(GgtdModel model)
        {
            try
            {
                var exists = await _dbContext.TblMdGiaGiaoTapDoanList.Where(x => x.Code == model.GgtdlHeader.Code).Select(x => x.IsActive).FirstOrDefaultAsync();
                
                if (exists == false)
                {
                    model.oldHeaderGgtd = "false";
                    return model;
                }
                else
                {
                    _dbContext.TblMdGiaGiaoTapDoanList.Update(model.GgtdlHeader);
                    _dbContext.TblMdGiaGiaoTapDoan.UpdateRange(model.Ggtd);

                    await _dbContext.SaveChangesAsync();
                    return model;
                   
                }
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
                var query = _dbContext.TblMdGiaGiaoTapDoan.AsQueryable();
                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Code.Contains(filter.KeyWord));
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

    public class GgtdModel
    {
        public string? oldHeaderGgtd { set; get; } 
        public TblMdGiaGiaoTapDoanList? GgtdlHeader { get; set; } = new TblMdGiaGiaoTapDoanList();
        public List<TblMdGiaGiaoTapDoan?> Ggtd { get; set; } = new List<TblMdGiaGiaoTapDoan>();

    }
}


