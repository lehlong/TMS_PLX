using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.EntityFrameworkCore;
using Microsoft.SqlServer.Server;
using NPOI.HSSF.Record.Chart;
using NPOI.OpenXml4Net.Util;

namespace DMS.BUSINESS.Services.BU
{
    public interface ICalculateDiscountService : IGenericService<TblBuCalculateDiscount, CalculateDiscountDto>
    {
        Task<CalculateDiscountInputModel> GenarateCreate();
        Task<CalculateDiscountInputModel> GetInput(string id);
        Task UpdateInput(CalculateDiscountInputModel input);
        Task Create(CalculateDiscountInputModel input);
        Task<CalculateDiscountOutputModel> CalculateDiscountOutput(string id);
    }
    public class CalculateDiscountService(AppDbContext dbContext, IMapper mapper) : GenericService<TblBuCalculateDiscount, CalculateDiscountDto>(dbContext, mapper), ICalculateDiscountService
    {

        #region Tìm kiếm các đợt nhập
        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var query = _dbContext.TblBuCalculateDiscount.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x =>
                    x.Name.Contains(filter.KeyWord));
                }
                return await Paging(query.OrderByDescending(x => x.CreateDate), filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        #endregion

        #region Tạo trước các thông tin khi ấn Tạo mới
        public async Task<CalculateDiscountInputModel> GenarateCreate()
        {
            try
            {
                var headerId = Guid.NewGuid().ToString();
                var lstGoods = await _dbContext.TblMdGoods.OrderBy(x => x.CreateDate).OrderBy(x => x.Order).ToListAsync();
                var lstMarket = await _dbContext.TblMdMarket.OrderBy(x => x.Code).ToListAsync();
                var lstCustomerDb = await _dbContext.TblMdCustomerDb.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerPt = await _dbContext.TblMdCustomerPt.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerFob = await _dbContext.TblMdCustomerFob.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerTnpp = await _dbContext.TblMdCustomerTnpp.OrderBy(x => x.Order).ToListAsync();


                return new CalculateDiscountInputModel
                {
                    Header = new TblBuCalculateDiscount
                    {
                        Id = headerId,
                        Date = DateTime.Now,
                        IsActive = true,
                        Status = "00"
                    },
                    InputPrice = lstGoods.Select(g => new TblBuInputPrice
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        GoodCode = g.Code,
                        GoodName = g.Name,
                        ThueBvmt = g.ThueBvmt,
                        Vcf = g.Vfc,
                        ChenhLech = g.MtsV1,
                        GblV1 = 0,
                        GblV2 = 0,
                        L15Blv2 = 0,
                        L15Nbl = 0,
                        LaiGop = 0,
                        IsActive = true,
                        Order = g.Order,
                    }).ToList(),
                    Market = lstMarket.Select(x => new TblBuInputMarket
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        FullName = x.FullName,
                        Local2 = x.Local2,
                        LocalCode = x.LocalCode,
                        WarehouseCode = x.WarehouseCode,
                        Gap = x.Gap,
                        Coefficient = x.Coefficient,
                        CuocVCBQ = x.CuocVCBQ,
                        CPChungChuaCuocVC = x.CPChungChuaCuocVC,
                        CkDieuTietDau = x.CkDieuTietDau,
                        CkDieuTietXang = x.CkDieuTietXang,
                    }).ToList(),

                    CustomerDb = lstCustomerDb.Select(x => new TblBuInputCustomerDb
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        LocalCode = x.LocalCode,
                        Local2 = x.Local2,
                        MarketCode = x.MarketCode,
                        CuLyBq = x.CuLyBq,
                        Cpccvc = x.Cpccvc,
                        Cvcbq = x.Cvcbq,
                        Lvnh = x.Lvnh,
                        Htcvc = x.Htcvc,
                        HttVb1370 = x.HttVb1370,
                        Ckv2 = x.Ckv2,
                        PhuongThuc = x.PhuongThuc,
                        Thtt = x.Thtt,
                        Order = x.Order,
                        Adrress = x.Adrress,
                        CkDau = x.CkDau,
                        CkXang = x.CkXang,
                        IsActive = true
                    }).ToList(),
                    CustomerPt = lstCustomerPt.Select(x => new TblBuInputCustomerPt
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        LocalCode = x.LocalCode,
                        MarketCode = x.MarketCode,
                        CuLyBq = x.CuLyBq,
                        Cpccvc = x.Cpccvc,
                        Cvcbq = x.Cvcbq,
                        Lvnh = x.Lvnh,
                        Htcvc = x.Htcvc,
                        HttVb1370 = x.HttVb1370,
                        Ckv2 = x.Ckv2,
                        PhuongThuc = x.PhuongThuc,
                        Thtt = x.Thtt,
                        Order = x.Order,
                        Adrress = x.Adrress,
                        CkDau = x.CkDau,
                        CkXang = x.CkXang,
                        IsActive = true
                    }).ToList(),
                    CustomerFob = lstCustomerFob.Select(x => new TblBuInputCustomerFob
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        LocalCode = x.LocalCode,
                        MarketCode = x.MarketCode,
                        CuLyBq = x.CuLyBq,
                        Cpccvc = x.Cpccvc,
                        Cvcbq = x.Cvcbq,
                        Lvnh = x.Lvnh,
                        Htcvc = x.Htcvc,
                        HttVb1370 = x.HttVb1370,
                        Ckv2 = x.Ckv2,
                        PhuongThuc = x.PhuongThuc,
                        Thtt = x.Thtt,
                        Order = x.Order,
                        Adrress = x.Adrress,
                        CkDau = x.CkDau,
                        CkXang = x.CkXang,
                        IsActive = true
                    }).ToList(),
                    CustomerTnpp = lstCustomerTnpp.Select(x => new TblBuInputCustomerTnpp
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        LocalCode = x.LocalCode,
                        MarketCode = x.MarketCode,
                        CuLyBq = x.CuLyBq,
                        Cpccvc = x.Cpccvc,
                        Cvcbq = x.Cvcbq,
                        Lvnh = x.Lvnh,
                        Htcvc = x.Htcvc,
                        HttVb1370 = x.HttVb1370,
                        Ckv2 = x.Ckv2,
                        PhuongThuc = x.PhuongThuc,
                        Thtt = x.Thtt,
                        Order = x.Order,
                        Adrress = x.Adrress,
                        CkDau = x.CkDau,
                        CkXang = x.CkXang,
                        IsActive = true
                    }).ToList()
                };

            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return new CalculateDiscountInputModel();
            }
        }
        #endregion

        #region Tạo các thông tin đầu vào
        public async Task Create(CalculateDiscountInputModel input)
        {
            try
            {
                _dbContext.TblBuCalculateDiscount.Add(input.Header);
                _dbContext.TblBuInputPrice.AddRange(input.InputPrice);
                _dbContext.TblBuInputMarket.AddRange(input.Market);
                _dbContext.TblBuInputCustomerDb.AddRange(input.CustomerDb);
                _dbContext.TblBuInputCustomerPt.AddRange(input.CustomerPt);
                _dbContext.TblBuInputCustomerFob.AddRange(input.CustomerFob);
                _dbContext.TblBuInputCustomerTnpp.AddRange(input.CustomerTnpp);
                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }
        #endregion

        #region Lấy thông tin đầu vào để cập nhật
        public async Task<CalculateDiscountInputModel> GetInput(string id)
        {
            try
            {
                return new CalculateDiscountInputModel
                {
                    Header = _dbContext.TblBuCalculateDiscount.Find(id),
                    InputPrice = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    Market = await _dbContext.TblBuInputMarket.Where(x => x.HeaderId == id).OrderBy(x => x.Code).ToListAsync(),
                    CustomerDb = await _dbContext.TblBuInputCustomerDb.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    CustomerPt = await _dbContext.TblBuInputCustomerPt.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    CustomerFob = await _dbContext.TblBuInputCustomerFob.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    CustomerTnpp = await _dbContext.TblBuInputCustomerTnpp.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                };
            }
            catch (Exception ex)
            {
                return new CalculateDiscountInputModel();
            }
        }
        #endregion

        #region Cập nhật lại dữ liệu đầu vào
        public async Task UpdateInput(CalculateDiscountInputModel input)
        {
            try
            {
                _dbContext.TblBuCalculateDiscount.Update(input.Header);
                _dbContext.TblBuInputPrice.UpdateRange(input.InputPrice);
                _dbContext.TblBuInputMarket.UpdateRange(input.Market);
                _dbContext.TblBuInputCustomerDb.UpdateRange(input.CustomerDb);
                _dbContext.TblBuInputCustomerPt.UpdateRange(input.CustomerPt);
                _dbContext.TblBuInputCustomerFob.UpdateRange(input.CustomerFob);
                _dbContext.TblBuInputCustomerTnpp.UpdateRange(input.CustomerTnpp);
                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }
        #endregion

        #region Tính toán đầu ra
        public async Task<CalculateDiscountOutputModel> CalculateDiscountOutput(string id)
        {
            try
            {
                #region Lấy các dữ liệu từ database
                var data = new CalculateDiscountOutputModel();
                var lstLocal = await _dbContext.tblMdLocal.OrderBy(x => x.Code).ToListAsync();
                var lstGoods = await _dbContext.TblMdGoods.OrderBy(x => x.Order).ToListAsync();
                var lstMarket = await _dbContext.TblBuInputMarket.Where(x => x.HeaderId == id).OrderBy(x => x.Code).ToListAsync();

                var lstCustomerDb = await _dbContext.TblBuInputCustomerDb.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerPt = await _dbContext.TblBuInputCustomerPt.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerFob = await _dbContext.TblBuInputCustomerFob.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerTnpp = await _dbContext.TblBuInputCustomerTnpp.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();

                var currentHeader = _dbContext.TblBuCalculateDiscount.Find(id);

                var currentData = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var previousHeader = await _dbContext.TblBuCalculateDiscount.Where(x => x.Date < currentHeader.Date).OrderByDescending(x => x.Date).FirstOrDefaultAsync();
                var previousData = new List<TblBuInputPrice>();
                if (previousHeader != null)
                {
                    previousData = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == previousHeader.Id).ToListAsync();
                }
                #endregion

                #region Dữ liệu gốc
                var _oDlg = 1;
                foreach (var i in currentData)
                {
                    data.Dlg.Dlg1.Add(new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.GblV1,
                        Col2 = i.GblV2,
                        Col3 = i.GblV2 - i.GblV1,
                    });
                    data.Dlg.Dlg2.Add(new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.GblV1 + i.ChenhLech,
                        Col2 = i.GblV2,
                        Col3 = i.GblV2 - i.ChenhLech - i.GblV1,
                    });
                    data.Dlg.Dlg4.Add(new DlgModel
                    {
                        Stt = _oDlg.ToString(),
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = i.L15Blv2,
                        Col4 = i.L15Blv2 * i.Vcf,
                        Col5 = (i.ThueBvmt + i.L15Blv2 * i.Vcf) * 1.1M,
                        Col6 = i.GblV1,
                        Col7 = i.GblV2,
                        Col8 = i.GblV2 / 1.1M - i.ThueBvmt,
                        Col9 = i.GblV2 - (i.ThueBvmt + i.L15Blv2 * i.Vcf) * 1.1M,
                        Col10 = (i.GblV2 / 1.1M - i.ThueBvmt) - (i.L15Blv2 * i.Vcf)
                    });
                    data.Dlg.Dlg5.Add(new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = previousData.Count() == 0 ? 0 : previousData.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.L15Blv2),
                        Col2 = i.L15Blv2,
                        Col3 = previousData.Count() == 0 ? i.L15Blv2 : i.L15Blv2 - previousData.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.L15Blv2),
                    });
                    data.Dlg.Dlg7.Add(new DlgModel
                    {
                        Stt = _oDlg.ToString(),
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = i.L15Blv2,
                        Col4 = i.Vcf * i.L15Blv2,
                        Col5 = ((i.Vcf * i.L15Blv2) + i.ThueBvmt) * 1.1M,
                    });
                    data.Dlg.Dlg8.Add(new DlgModel
                    {
                        Stt = _oDlg.ToString(),
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = i.ThueBvmt / i.Vcf,
                        Col4 = i.L15Blv2,
                        Col5 = 0,
                        Col6 = (i.ThueBvmt / i.Vcf) + i.L15Blv2,
                        Col7 = ((i.ThueBvmt / i.Vcf) + i.L15Blv2) * 1.1M,
                        Note = ""
                    });
                    _oDlg++;
                }

                data.Dlg.Dlg3.Add(new DlgModel
                {
                    GoodName = "Vùng 1+ (TP Vinh)",
                    IsBold = true
                });
                foreach (var i in currentData)
                {
                    var d = new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        LocalCode = "V1",
                        Col1 = previousData.Count() == 0 ? 0 : previousData.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.GblV1 + x.ChenhLech),
                        Col2 = i.GblV1 + i.ChenhLech
                    };
                    d.Col3 = d.Col2 - d.Col1;
                    data.Dlg.Dlg3.Add(d);
                }
                data.Dlg.Dlg3.Add(new DlgModel
                {
                    GoodName = "Vùng 2 (các địa bàn còn lại)",
                    IsBold = true
                });
                foreach (var i in currentData)
                {
                    var d = new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        LocalCode = "V2",
                        Col1 = previousData.Count() == 0 ? 0 : previousData.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.GblV2),
                        Col2 = i.GblV2
                    };
                    d.Col3 = d.Col2 - d.Col1;
                    data.Dlg.Dlg3.Add(d);
                }


                data.Dlg.Dlg6.Add(new DlgModel
                {
                    GoodName = "Vùng thị trường trung tâm",
                    Stt = "I",
                    IsBold = true
                });
                var _o1 = 1;
                foreach (var i in currentData)
                {
                    data.Dlg.Dlg6.Add(new DlgModel
                    {
                        Stt = _o1.ToString(),
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        LocalCode = "V1",
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = i.L15Nbl,
                        Col4 = i.Vcf * i.L15Nbl,
                        Col5 = (i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M,
                        Col6 = data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1),
                        Col7 = data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) / 1.1M - i.ThueBvmt,
                        Col8 = data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M),
                        Col9 = (data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M,
                        Col10 = i.LaiGop,
                        Col11 = i.Vcf * i.LaiGop * 1.1M,
                        Col12 = ((i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M)) * 1.1M,
                        Col13 = (i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M),
                        Col14 = i.FobV1,
                        Col15 = ((((i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M)) * 1.1M) - i.FobV1) * i.Vcf,
                        Col16 = (((i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col1) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M)) * 1.1M) - i.FobV1
                    });
                    _o1++;
                }
                data.Dlg.Dlg6.Add(new DlgModel
                {
                    GoodName = "Các vùng thị trường còn lại",
                    Stt = "II",
                    IsBold = true
                });
                var _o2 = 1;
                foreach (var i in currentData)
                {
                    data.Dlg.Dlg6.Add(new DlgModel
                    {
                        Stt = _o2.ToString(),
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        LocalCode = "V2",
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = i.L15Nbl,
                        Col4 = i.Vcf * i.L15Nbl,
                        Col5 = (i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M,
                        Col6 = data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2),
                        Col7 = data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) / 1.1M - i.ThueBvmt,
                        Col8 = data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M),
                        Col9 = (data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M,
                        Col10 = i.LaiGop,
                        Col11 = i.Vcf * i.LaiGop * 1.1M,
                        Col12 = ((i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M)) * 1.1M,
                        Col13 = (i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M),
                        Col14 = i.FobV2,
                        Col15 = ((((i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M)) * 1.1M) - i.FobV2) * i.Vcf,
                        Col16 = (((i.Vcf * i.LaiGop * 1.1M) + ((data.Dlg.Dlg2.Where(x => x.GoodCode == i.GoodCode).Sum(x => x.Col2) - ((i.Vcf * i.L15Nbl + i.ThueBvmt) * 1.1M)) / 1.1M)) * 1.1M) - i.FobV2
                    });
                    _o2++;
                }
                #endregion

                #region PT
                foreach (var l in lstLocal)
                {
                    data.Pt.Add(new DataModel
                    {
                        MarketName = l.Name,
                        IsBold = true,
                    });
                    var market = lstMarket.Where(x => x.LocalCode == l.Code).OrderBy(x => x.Code).ToList();
                    var _o = 1;
                    foreach (var i in market)
                    {
                        var p = new DataModel
                        {
                            Stt = _o.ToString(),
                            MarketCode = i.Code,
                            MarketName = i.Name,
                            Col1 = i.Gap ?? 0,
                            Col2 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col13),
                            Col3 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col13),
                            Col4 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col13),
                            Col5 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col13),
                            Col6 = i.CPChungChuaCuocVC + i.CuocVCBQ ?? 0,
                            Col7 = i.CPChungChuaCuocVC ?? 0,
                            Col8 = i.CuocVCBQ ?? 0,
                            Col9 = i.CkDieuTietXang ?? 0,
                            Col10 = i.CkDieuTietDau ?? 0,
                            Col11 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietXang ?? 0,
                            Col13 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietXang ?? 0,
                            Col15 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietDau ?? 0,
                            Col17 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietDau ?? 0,

                        };
                        p.Col12 = Math.Round(p.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;
                        p.Col14 = Math.Round(p.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;
                        p.Col16 = Math.Round(p.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;
                        p.Col18 = Math.Round(p.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;

                        p.Col19 = p.Col2 - p.Col6 - p.Col12;
                        p.Col20 = p.Col3 - p.Col6 - p.Col14;
                        p.Col21 = p.Col4 - p.Col6 - p.Col16;
                        p.Col22 = p.Col5 - p.Col6 - p.Col18;

                        p.Col24 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col6) - Math.Round(p.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10;
                        p.Col26 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col6) - Math.Round(p.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10; ;
                        p.Col28 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col6) - Math.Round(p.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10; ;
                        p.Col30 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col6) - Math.Round(p.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10; ;

                        p.Col23 = p.Col24 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col2);
                        p.Col25 = p.Col26 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col2);
                        p.Col27 = p.Col28 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col2);
                        p.Col29 = p.Col30 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col2);

                        data.Pt.Add(p);
                        _o++;
                    }
                }
                #endregion

                #region ĐB
                foreach (var l in lstLocal)
                {
                    data.Db.Add(new DataModel
                    {
                        CustomerName = l.Name,
                        IsBold = true,
                    });
                    var customer = lstCustomerDb.Where(x => x.LocalCode == l.Code).ToList();
                    var _o = 1;
                    foreach (var i in customer)
                    {
                        decimal col14, col16, col18, col20;

                        string marketCode = i.MarketCode;

                        if (i.Code == "318337" || i.Code == "322004")
                        {
                            marketCode = "V1_02";
                        }
                        else if (i.Code == "327403" || i.Code == "320754" || i.Code == "326869" || i.Code == "305023")
                        {
                            marketCode = "V2_03";
                        }

                        col14 = data.Pt.Where(x => x.MarketCode == marketCode)
                                   .Sum(x => Math.Round(x.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                        col16 = data.Pt.Where(x => x.MarketCode == marketCode)
                                   .Sum(x => Math.Round(x.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10);

                        if (i.Code == "906962")
                        {
                            col18 = data.Pt.Where(x => x.MarketCode == "V2_02")
                                     .Sum(x => Math.Round(x.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                            col20 = data.Pt.Where(x => x.MarketCode == "V2_02")
                                     .Sum(x => Math.Round(x.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                        }
                        else
                        {
                            col18 = data.Pt.Where(x => x.MarketCode == marketCode)
                                     .Sum(x => Math.Round(x.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                            col20 = data.Pt.Where(x => x.MarketCode == marketCode)
                                     .Sum(x => Math.Round(x.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                        }

                        var p = new DataModel
                        {
                            Stt = _o.ToString(),
                            CustomerCode = i.Code,
                            CustomerName = i.Name,
                            MarketCode = i.MarketCode,
                            MarketName = lstMarket.FirstOrDefault(x => x.Code == i.MarketCode)?.Name ?? "",
                            Col1 = i.CuLyBq,
                            Col2 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col13),
                            Col3 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col13),
                            Col4 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col13),
                            Col5 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col13),
                            Col6 = i.Cpccvc + i.Cvcbq + i.Lvnh,
                            Col7 = i.Cpccvc,
                            Col8 = i.Cvcbq,
                            Col9 = i.Lvnh,
                            Col10 = i.Htcvc,
                            Col11 = i.HttVb1370,
                            Col12 = i.CkXang,
                            Col13 = i.CkDau,
                            Col14 = i.HttVb1370 + i.CkXang + col14,
                            Col16 = i.HttVb1370 + i.CkXang + col16,
                            Col18 = i.HttVb1370 + col18,
                            Col20 = i.HttVb1370 + col20 + i.CkDau,
                        };
                        p.Col15 = p.Col14 / 1.1M;
                        p.Col17 = p.Col16 / 1.1M;
                        p.Col19 = p.Col18 / 1.1M;
                        p.Col21 = p.Col20 / 1.1M;

                        p.Col23 = p.Col2 - p.Col6 - p.Col15 - p.Col22 / 1.1M + p.Col11 / 1.1M;
                        p.Col24 = p.Col3 - p.Col6 - p.Col17 - p.Col22 / 1.1M + p.Col11 / 1.1M;
                        p.Col25 = p.Col4 - p.Col6 - p.Col19 - p.Col22 / 1.1M + p.Col11 / 1.1M;
                        p.Col26 = p.Col5 - p.Col6 - p.Col21 - p.Col22 / 1.1M + p.Col11 / 1.1M;

                        p.Col28 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col6) - p.Col14;
                        p.Col30 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col6) - p.Col16;
                        p.Col32 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col6) - p.Col18;
                        p.Col34 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col6) - p.Col20;


                        p.Col27 = p.Col28 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col2);
                        p.Col29 = p.Col30 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col2);
                        p.Col31 = p.Col32 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col2);
                        p.Col33 = p.Col34 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col2);


                        data.Db.Add(p);
                        _o++;
                    }
                }
                #endregion

                #region FOB
                foreach (var l in lstLocal)
                {
                    data.Fob.Add(new DataModel
                    {
                        CustomerName = l.Name,
                        IsBold = true,
                    });
                    var customer = lstCustomerFob.Where(x => x.LocalCode == l.Code).ToList();
                    var _o = 1;
                    foreach (var i in customer)
                    {
                        var p = new DataModel
                        {
                            Stt = _o.ToString(),
                            CustomerCode = i.Code,
                            CustomerName = i.Name,
                            Col1 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col13),
                            Col2 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col13),
                            Col3 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col13),
                            Col4 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col13),
                            Col5 = i.Cpccvc + i.Cvcbq + i.Lvnh,
                            Col6 = i.Cpccvc,
                            Col7 = i.Cvcbq,
                            Col8 = i.Lvnh,
                            Col9 = i.HttVb1370,
                            Col10 = i.CkXang,
                            Col11 = i.CkDau,
                            Col12 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col14) + i.CkXang + i.HttVb1370,
                            Col13 = (data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col14) + i.CkXang + i.HttVb1370) / 1.1M,
                            Col14 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col14) + i.CkXang + i.HttVb1370,
                            Col15 = (data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col14) + i.CkXang + i.HttVb1370) / 1.1M,
                            Col16 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col14) + i.CkDau + i.HttVb1370,
                            Col17 = (data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col14) + i.CkDau + i.HttVb1370) / 1.1M,
                            Col18 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col14) + i.CkDau + i.HttVb1370,
                            Col19 = (data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col14) + i.CkDau + i.HttVb1370) / 1.1M,
                            Col20 = i.Ckv2
                        };
                        p.Col21 = p.Col1 - p.Col5 - p.Col13 - p.Col20 / 1.1M + p.Col9 / 1.1M;
                        p.Col22 = p.Col2 - p.Col5 - p.Col15 - p.Col20 / 1.1M + p.Col9 / 1.1M;
                        p.Col23 = p.Col3 - p.Col5 - p.Col17 - p.Col20 / 1.1M + p.Col9 / 1.1M;
                        p.Col24 = p.Col4 - p.Col5 - p.Col19 - p.Col20 / 1.1M + p.Col9 / 1.1M;

                        p.Col26 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col6) - p.Col12;
                        p.Col28 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col6) - p.Col14;
                        p.Col30 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col6) - p.Col16;
                        p.Col32 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col6) - p.Col18;

                        p.Col25 = p.Col26 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201032").Sum(x => x.Col2);
                        p.Col27 = p.Col28 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0201004").Sum(x => x.Col2);
                        p.Col29 = p.Col30 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601005").Sum(x => x.Col2);
                        p.Col31 = p.Col32 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "0601002").Sum(x => x.Col2);


                        data.Fob.Add(p);
                        _o++;
                    }
                }
                #endregion

                #region PT09
                var _pt09 = 1;
                foreach (var i in lstCustomerTnpp)
                {
                    var p = new DataModel
                    {
                        Stt = _pt09.ToString(),
                        CustomerCode = i.Code,
                        CustomerName = i.Name,
                        Col1 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201032").Sum(x => x.Col13),
                        Col2 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201004").Sum(x => x.Col13),
                        Col3 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601005").Sum(x => x.Col13),
                        Col4 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601002").Sum(x => x.Col13),
                        Col5 = i.Cpccvc + i.Lvnh,
                        Col6 = i.Cpccvc,
                        Col7 = 0,
                        Col8 = i.Lvnh,
                        Col9 = i.HttVb1370,
                        Col10 = i.CkXang,
                        Col11 = i.CkDau,
                        Col12 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201032").Sum(x => x.Col14) + i.CkXang + i.HttVb1370,
                        Col13 = (data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201032").Sum(x => x.Col14) + i.CkXang + i.HttVb1370) / 1.1M,
                        Col14 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201004").Sum(x => x.Col14) + i.CkXang + i.HttVb1370,
                        Col15 = (data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201004").Sum(x => x.Col14) + i.CkXang + i.HttVb1370) / 1.1M,
                        Col16 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601005").Sum(x => x.Col14) + i.CkDau + i.HttVb1370,
                        Col17 = (data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601005").Sum(x => x.Col14) + i.CkDau + i.HttVb1370) / 1.1M,
                        Col18 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601002").Sum(x => x.Col14) + i.CkDau + i.HttVb1370,
                        Col19 = (data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601002").Sum(x => x.Col14) + i.CkDau + i.HttVb1370) / 1.1M,
                        Col20 = i.Ckv2
                    };
                    p.Col21 = p.Col1 - p.Col5 - p.Col13 - p.Col20 / 1.1M + p.Col9 / 1.1M;
                    p.Col22 = p.Col2 - p.Col5 - p.Col15 - p.Col20 / 1.1M + p.Col9 / 1.1M;
                    p.Col23 = p.Col3 - p.Col5 - p.Col17 - p.Col20 / 1.1M + p.Col9 / 1.1M;
                    p.Col24 = p.Col4 - p.Col5 - p.Col19 - p.Col20 / 1.1M + p.Col9 / 1.1M;

                    p.Col26 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201032").Sum(x => x.Col6) - p.Col12;
                    p.Col28 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201004").Sum(x => x.Col6) - p.Col14;
                    p.Col30 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601005").Sum(x => x.Col6) - p.Col16;
                    p.Col32 = data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601002").Sum(x => x.Col6) - p.Col18;

                    p.Col25 = p.Col26 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201032").Sum(x => x.Col2);
                    p.Col27 = p.Col28 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0201004").Sum(x => x.Col2);
                    p.Col29 = p.Col30 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601005").Sum(x => x.Col2);
                    p.Col31 = p.Col32 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == "V2" && x.GoodCode == "0601002").Sum(x => x.Col2);


                    data.Pt09.Add(p);
                    _pt09++;
                }
                #endregion

                #region PL1
                var local = lstMarket.Where(x => !string.IsNullOrEmpty(x.Local2)).OrderBy(x => x.Local2).Select(x => x.Local2).Distinct().ToList();
                var _pl1 = 1;
                var _pl1_lm = 1;
                foreach (var l in local)
                {
                    data.Pl1.Add(new DataModel
                    {
                        Stt = IntToRoman(_pl1_lm),
                        MarketName = l,
                        IsBold = true,
                    });
                    var market = lstMarket.Where(x => x.Local2 == l).ToList();
                    foreach(var i in market)
                    {
                        data.Pl1.Add(new DataModel
                        {
                            Stt = _pl1.ToString(),
                            MarketCode = i.Code,
                            MarketName = i.FullName,
                            Col1 = data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => Math.Round(x.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                            Col2 = data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => Math.Round(x.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                            Col3 = data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => Math.Round(x.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                            Col4 = data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => Math.Round(x.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                        });
                        _pl1++;
                    }
                    _pl1_lm++;
                }
                #endregion

                #region PL3
                var pl3 = 1;
                foreach(var i in lstCustomerFob)
                {
                    data.Pl3.Add(new DataModel
                    {
                        Stt = pl3.ToString(),
                        CustomerCode = i.Code,
                        CustomerName = i.Name,
                        Col1 = data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col12),
                        Col2 = data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col14),
                        Col3 = data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col16),
                        Col4 = data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col18),

                    });

                    pl3++;
                }
                #endregion

                #region PL4
                var pl4 = 1;
                foreach (var i in lstCustomerTnpp)
                {
                    data.Pl4.Add(new DataModel
                    {
                        Stt = pl4.ToString(),
                        CustomerCode = i.Code,
                        CustomerName = i.Name,
                        Col1 = data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col12),
                        Col2 = data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col14),
                        Col3 = data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col16),
                        Col4 = data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col18),
                    });

                    pl4++;
                }
                #endregion

                #region VK11-PT

                foreach (var g in lstGoods)
                {
                    data.Vk11Pt.Add(new VK11Model
                    {
                        CustomerName = g.Name,
                        IsBold = true,
                    });

                    var o = 1;
                    foreach (var i in lstCustomerPt)
                    {
                        var value = g.Code == "0201032" ? data.Pt.Where(x => x.MarketCode == i.MarketCode).Sum(x => x.Col23)
                            : g.Code == "0201004" ? data.Pt.Where(x => x.MarketCode == i.MarketCode).Sum(x => x.Col25)
                            : g.Code == "0601005" ? data.Pt.Where(x => x.MarketCode == i.MarketCode).Sum(x => x.Col27)
                            : data.Pt.Where(x => x.MarketCode == i.MarketCode).Sum(x => x.Col29);
                        data.Vk11Pt.Add(new VK11Model
                        {
                            Stt = o.ToString(),
                            CustomerName = i.Name,
                            Address = i.Adrress,
                            MarketCode = i.MarketCode,
                            MarketName = lstMarket.FirstOrDefault(x => x.Code == i.MarketCode)?.Name ?? "",
                            Col1 = i.CuLyBq,
                            Col2 = i.Cvcbq,
                            Col3 = i.PhuongThuc,
                            Col4 = i.Code,
                            Col5 = g.Code,
                            Col7 = i.Thtt,
                            Col8 = Math.Round(value),
                            Col13 = currentHeader.Date.ToString("dd.MM.yyyy"),
                            Col14 = currentHeader.Date.ToString("hh:mm"),
                        });
                        o++;
                    }
                }

                #endregion

                #region VK11-ĐB
                foreach (var g in lstGoods)
                {
                    data.Vk11Db.Add(new VK11Model
                    {
                        CustomerName = g.Name,
                        IsBold = true,
                    });

                    var o = 1;
                    foreach (var i in lstCustomerDb)
                    {
                        var value = g.Code == "0201032" ? data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col27)
                            : g.Code == "0201004" ? data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col29)
                            : g.Code == "0601005" ? data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col31)
                            : data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col33);
                        data.Vk11Db.Add(new VK11Model
                        {
                            Stt = o.ToString(),
                            CustomerName = i.Name,
                            Address = i.Adrress,
                            MarketCode = i.MarketCode,
                            MarketName = lstMarket.FirstOrDefault(x => x.Code == i.MarketCode)?.Name ?? "",
                            Col1 = i.CuLyBq,
                            Col2 = i.Cvcbq,
                            Col3 = i.PhuongThuc,
                            Col4 = i.Code,
                            Col5 = g.Code,
                            Col7 = i.Thtt,
                            Col8 = Math.Round(value),
                            Col13 = currentHeader.Date.ToString("dd.MM.yyyy"),
                            Col14 = currentHeader.Date.ToString("hh:mm"),
                        });
                        o++;
                    }
                }
                #endregion

                #region VK11-FOB
                foreach (var g in lstGoods)
                {
                    data.Vk11Fob.Add(new VK11Model
                    {
                        CustomerName = g.Name,
                        IsBold = true,
                    });

                    var o = 1;
                    foreach (var i in lstCustomerFob)
                    {
                        var value = g.Code == "0201032" ? data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col25)
                            : g.Code == "0201004" ? data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col27)
                            : g.Code == "0601005" ? data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col29)
                            : data.Fob.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col31);
                        data.Vk11Fob.Add(new VK11Model
                        {
                            Stt = o.ToString(),
                            CustomerName = i.Name,
                            Address = i.Adrress,
                            MarketCode = i.MarketCode,
                            MarketName = lstMarket.FirstOrDefault(x => x.Code == i.MarketCode)?.Name ?? "",
                            Col1 = i.CuLyBq,
                            Col2 = i.Cvcbq,
                            Col3 = i.PhuongThuc,
                            Col4 = i.Code,
                            Col5 = g.Code,
                            Col7 = i.Thtt,
                            Col8 = Math.Round(value),
                            Col13 = currentHeader.Date.ToString("dd.MM.yyyy"),
                            Col14 = currentHeader.Date.ToString("hh:mm"),
                        });
                        o++;
                    }
                }
                #endregion

                #region VK11-TNPP
                foreach (var g in lstGoods)
                {
                    data.Vk11Tnpp.Add(new VK11Model
                    {
                        CustomerName = g.Name,
                        IsBold = true,
                    });

                    var o = 1;
                    foreach (var i in lstCustomerTnpp)
                    {
                        var value = g.Code == "0201032" ? data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col25)
                            : g.Code == "0201004" ? data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col27)
                            : g.Code == "0601005" ? data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col29)
                            : data.Pt09.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col31);
                        data.Vk11Tnpp.Add(new VK11Model
                        {
                            Stt = o.ToString(),
                            CustomerName = i.Name,
                            Address = i.Adrress,
                            MarketCode = i.MarketCode,
                            MarketName = lstMarket.FirstOrDefault(x => x.Code == i.MarketCode)?.Name ?? "",
                            Col1 = i.CuLyBq,
                            Col2 = i.Cvcbq,
                            Col3 = i.PhuongThuc,
                            Col4 = i.Code,
                            Col5 = g.Code,
                            Col7 = i.Thtt,
                            Col8 = Math.Round(value),
                            Col13 = currentHeader.Date.ToString("dd.MM.yyyy"),
                            Col14 = currentHeader.Date.ToString("hh:mm"),
                        });
                        o++;
                    }
                }
                #endregion

                return RoundNumber(data);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return new CalculateDiscountOutputModel();
            }
        }
        #endregion

        #region Làm tròn dữ liệu
        public CalculateDiscountOutputModel RoundNumber(CalculateDiscountOutputModel data)
        {
            #region DLG
            foreach (var i in data.Dlg.Dlg4)
            {
                i.Col4 = Math.Round(i.Col4);
                i.Col5 = Math.Round(i.Col5);
                i.Col8 = Math.Round(i.Col8);
                i.Col9 = Math.Round(i.Col9);
                i.Col10 = Math.Round(i.Col10);
            }
            foreach (var i in data.Dlg.Dlg5)
            {
                i.Col1 = Math.Round(i.Col1);
                i.Col3 = Math.Round(i.Col3);
            }
            foreach (var i in data.Dlg.Dlg7)
            {
                i.Col4 = Math.Round(i.Col4);
                i.Col5 = Math.Round(i.Col5 / 10, 0, MidpointRounding.AwayFromZero) * 10;
            }
            foreach (var i in data.Dlg.Dlg8)
            {
                i.Col3 = Math.Round(i.Col1);
                i.Col6 = Math.Round(i.Col6);
                i.Col7 = Math.Round(i.Col7);
            }
            foreach (var i in data.Dlg.Dlg3)
            {
                i.Col1 = Math.Round(i.Col1);
                i.Col2 = Math.Round(i.Col2);
                i.Col3 = Math.Round(i.Col3);
            }
            foreach (var i in data.Dlg.Dlg6)
            {
                i.Col4 = Math.Round(i.Col4);
                i.Col5 = Math.Round(i.Col5);
                i.Col7 = Math.Round(i.Col7);
                i.Col8 = Math.Round(i.Col8);
                i.Col9 = Math.Round(i.Col9);
                i.Col11 = Math.Round(i.Col11);
                i.Col12 = Math.Round(i.Col12);
                i.Col13 = Math.Round(i.Col13);
                i.Col15 = Math.Round(i.Col15);
                i.Col16 = Math.Round(i.Col16);
            }
            #endregion

            #region PT
            foreach (var i in data.Pt)
            {
                i.Col2 = Math.Round(i.Col2);
                i.Col3 = Math.Round(i.Col3);
                i.Col4 = Math.Round(i.Col4);
                i.Col5 = Math.Round(i.Col5);
                i.Col11 = Math.Round(i.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10;
                i.Col13 = Math.Round(i.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10;
                i.Col15 = Math.Round(i.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10;
                i.Col17 = Math.Round(i.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10;
                i.Col14 = Math.Round(i.Col14);
                i.Col16 = Math.Round(i.Col16);
                i.Col18 = Math.Round(i.Col18);
                i.Col12 = Math.Round(i.Col12);
                i.Col19 = Math.Round(i.Col19);
                i.Col20 = Math.Round(i.Col20);
                i.Col21 = Math.Round(i.Col21);
                i.Col22 = Math.Round(i.Col22);
                i.Col23 = Math.Round(i.Col23);
                i.Col25 = Math.Round(i.Col25);
                i.Col27 = Math.Round(i.Col27);
                i.Col29 = Math.Round(i.Col29);
            }
            #endregion

            #region ĐB
            foreach (var i in data.Db)
            {
                i.Col2 = Math.Round(i.Col2);
                i.Col3 = Math.Round(i.Col3);
                i.Col4 = Math.Round(i.Col4);
                i.Col5 = Math.Round(i.Col5);
                i.Col15 = Math.Round(i.Col15);
                i.Col17 = Math.Round(i.Col17);
                i.Col19 = Math.Round(i.Col19);
                i.Col21 = Math.Round(i.Col21);
                i.Col23 = Math.Round(i.Col23);
                i.Col24 = Math.Round(i.Col24);
                i.Col25 = Math.Round(i.Col25);
                i.Col26 = Math.Round(i.Col26);
                i.Col27 = Math.Round(i.Col27);
                i.Col29 = Math.Round(i.Col29);
                i.Col31 = Math.Round(i.Col31);
                i.Col33 = Math.Round(i.Col33);
            }
            #endregion

            #region FOB
            foreach (var i in data.Fob)
            {
                i.Col1 = Math.Round(i.Col1);
                i.Col2 = Math.Round(i.Col2);
                i.Col3 = Math.Round(i.Col3);
                i.Col4 = Math.Round(i.Col4);

                i.Col13 = Math.Round(i.Col13);
                i.Col15 = Math.Round(i.Col15);
                i.Col17 = Math.Round(i.Col17);
                i.Col19 = Math.Round(i.Col19);

                i.Col21 = Math.Round(i.Col21);
                i.Col22 = Math.Round(i.Col22);
                i.Col23 = Math.Round(i.Col23);
                i.Col24 = Math.Round(i.Col24);
                i.Col25 = Math.Round(i.Col25);
                i.Col26 = Math.Round(i.Col26);
                i.Col27 = Math.Round(i.Col27);
                i.Col28 = Math.Round(i.Col28);
                i.Col29 = Math.Round(i.Col29);
                i.Col30 = Math.Round(i.Col30);
                i.Col31 = Math.Round(i.Col31);
                i.Col32 = Math.Round(i.Col32);
            }
            #endregion

            #region PT09
            foreach (var i in data.Pt09)
            {
                i.Col1 = Math.Round(i.Col1);
                i.Col2 = Math.Round(i.Col2);
                i.Col3 = Math.Round(i.Col3);
                i.Col4 = Math.Round(i.Col4);

                i.Col13 = Math.Round(i.Col13);
                i.Col15 = Math.Round(i.Col15);
                i.Col17 = Math.Round(i.Col17);
                i.Col19 = Math.Round(i.Col19);

                i.Col21 = Math.Round(i.Col21);
                i.Col22 = Math.Round(i.Col22);
                i.Col23 = Math.Round(i.Col23);
                i.Col24 = Math.Round(i.Col24);
                i.Col25 = Math.Round(i.Col25);
                i.Col26 = Math.Round(i.Col26);
                i.Col27 = Math.Round(i.Col27);
                i.Col28 = Math.Round(i.Col28);
                i.Col29 = Math.Round(i.Col29);
                i.Col30 = Math.Round(i.Col30);
                i.Col31 = Math.Round(i.Col31);
                i.Col32 = Math.Round(i.Col32);
            }
            #endregion

            return data;
        }
        #endregion

        #region Chuyển số thành chữ la mã
        public string IntToRoman(int num)
        {
            var romanNumerals = new (int value, string numeral)[]
            {
            (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
            (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
            (10, "X"), (9, "IX"), (5, "V"), (4, "IV"),
            (1, "I")
            };

            string result = "";
            foreach (var (value, numeral) in romanNumerals)
            {
                while (num >= value)
                {
                    result += numeral;
                    num -= value;
                }
            }
            return result;
        }
        #endregion
    }
}



