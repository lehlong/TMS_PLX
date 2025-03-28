using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.EntityFrameworkCore;
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
                var data = new CalculateDiscountOutputModel();
                var lstLocal = await _dbContext.tblMdLocal.OrderBy(x => x.Code).ToListAsync();
                var currentHeader = _dbContext.TblBuCalculateDiscount.Find(id);
                var currentData = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var previousHeader = await _dbContext.TblBuCalculateDiscount.Where(x => x.Date < currentHeader.Date).OrderByDescending(x => x.Date).FirstOrDefaultAsync();
                var previousData = new List<TblBuInputPrice>();
                if (previousHeader != null)
                {
                    previousData = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == previousHeader.Id).ToListAsync();
                }

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
                foreach(var l in lstLocal)
                {
                    data.Pt.Add(new PtModel
                    {
                        MarketName = l.Name,
                        IsBold = true,
                    });
                    var market = await _dbContext.TblBuInputMarket.Where(x => x.HeaderId == id && x.LocalCode == l.Code).OrderBy(x => x.Code).ToListAsync();
                    var _o = 1;
                    foreach (var i in market)
                    {
                        var p = new PtModel
                        {
                            Stt = _o.ToString(),
                            MarketCode = i.Code,
                            MarketName = i.Name,
                            Col1 = i.Gap ?? 0,
                            Col2 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201032").Sum(x => x.Col13),
                            Col3 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201004").Sum(x => x.Col13),
                            Col4 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601005").Sum(x => x.Col13),
                            Col5 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601002").Sum(x => x.Col13),
                            Col6 = i.CPChungChuaCuocVC + i.CuocVCBQ ?? 0,
                            Col7 = i.CPChungChuaCuocVC ?? 0,
                            Col8 = i.CuocVCBQ ?? 0,
                            Col9 = i.CkDieuTietXang ?? 0,
                            Col10 = i.CkDieuTietDau ?? 0,
                            Col11 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201032").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietXang ?? 0,
                            Col13 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201004").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietXang ?? 0,
                            Col15 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601005").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietDau ?? 0,
                            Col17 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601002").Sum(x => x.Col14) - i.CuocVCBQ * 1.1M + i.CkDieuTietDau ?? 0,

                        };
                        p.Col12 = Math.Round(p.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;
                        p.Col14 = Math.Round(p.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;
                        p.Col16 = Math.Round(p.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;
                        p.Col18 = Math.Round(p.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10 / 1.1M;

                        p.Col19 = p.Col2 - p.Col6 - p.Col12;
                        p.Col20 = p.Col3 - p.Col6 - p.Col14;
                        p.Col21 = p.Col4 - p.Col6 - p.Col16;
                        p.Col22 = p.Col5 - p.Col6 - p.Col18;

                        p.Col24 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201032").Sum(x => x.Col6) - Math.Round(p.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10;
                        p.Col26 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201004").Sum(x => x.Col6) - Math.Round(p.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10; ;
                        p.Col28 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601005").Sum(x => x.Col6) - Math.Round(p.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10; ;
                        p.Col30 = data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601002").Sum(x => x.Col6) - Math.Round(p.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10; ;

                        p.Col23 = p.Col24 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201032").Sum(x => x.Col2);
                        p.Col25 = p.Col26 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "201004").Sum(x => x.Col2);
                        p.Col27 = p.Col28 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601005").Sum(x => x.Col2);
                        p.Col29 = p.Col30 / 1.1M - data.Dlg.Dlg6.Where(x => x.LocalCode == l.Code && x.GoodCode == "601002").Sum(x => x.Col2);

                        data.Pt.Add(p);
                        _o++;
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

            return data;
        }
        #endregion
    }
}



