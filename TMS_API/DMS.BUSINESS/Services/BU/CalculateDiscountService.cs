﻿using AutoMapper;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Extentions;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using DocumentFormat.OpenXml.Bibliography;
using Microsoft.EntityFrameworkCore;
using Microsoft.SqlServer.Server;
using NPOI.HSSF.Record.Chart;
using NPOI.OpenXml4Net.Util;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using PROJECT.Service.Extention;
using Aspose.Words;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;
using DMS.CORE.Entities.IN;
using NPOI.XSSF.UserModel.Helpers;
using static System.Runtime.InteropServices.JavaScript.JSType;
using DMS.BUSINESS.Dtos.MD;
using NPOI.SS.Formula.Functions;
using DMS.CORE.Entities.MD;
using Microsoft.IdentityModel.Tokens;
using DocumentFormat.OpenXml.Office.CustomUI;
using DocumentFormat.OpenXml.Office2010.CustomUI;

namespace DMS.BUSINESS.Services.BU
{
    public interface ICalculateDiscountService : IGenericService<TblBuCalculateDiscount, CalculateDiscountDto>
    {
        Task<CalculateDiscountInputModel> GenarateCreate();
        Task<CalculateDiscountInputModel> GetInput(string id);
        Task UpdateInput(CalculateDiscountInputModel input);
        Task Create(CalculateDiscountInputModel input);
        Task<List<TblBuHistoryAction>> GetHistoryAction(string code);
        Task HandleQuyTrinh(QuyTrinhModel data);
        Task<CalculateDiscountOutputModel> CalculateDiscountOutput(string id);
        Task<string> ExportExcel(string headerId);
        Task<string> ExportExcelTPKD(string headerId);
        Task<string> ExportExcelNVK(string headerId);
        Task<string> GenarateWordTrinhKy(string headerId, string nameTeam);
        Task<string> GenarateWord(List<CustomBBDOExportWord> lstCustomerChecked, string headerId);
        Task<List<string>> GenarateFile(List<string> lstCustomerChecked, string type, string headerId, CalculateDiscountInputModel data, List<CustomBBDOExportWord>? lstCustomerCheckedWord = null);
        Task  GenarateFileMail (List<string> lstCustomerChecked, string type, string headerId, List<CustomBBDOExportWord>? lstCustomerCheckedWord = null);
        Task<string> ExportExcelTrinhKy(string headerId);
        Task<List<TblBuHistoryDownload>> GetHistoryFile(string code);
        Task SendEmail(string headerId);
        Task SendlstMail(List<string> lstEmail);
        Task<string> SaveSMS(string headerId, string smsName);
        Task SendSMS(List<string> lstSms);
        Task<List<NotifyEmailViewModel>> GetHistoryMail(string headerId);
        Task ResetSendlstMail(List<string> lstEmail);
        Task ResetSendSMS(List<string> lstSms);
        //Task ResendMail(string headerId);
        Task<List<TblNotifySms>> GetHistorySms(string headerId);
        Task<List<TblBuInputCustomerBbdo>> GetCustomerBbdo(string id);
        Task<List<CustomInput>> GetAllInputCustomer(); 
        Task<CalculateDiscountInputModel> CopyInput(string headerId, string id);
    }
    public class CalculateDiscountService(AppDbContext dbContext, IMapper mapper) : GenericService<TblBuCalculateDiscount, CalculateDiscountDto>(dbContext, mapper), ICalculateDiscountService
    {
        private readonly IDiscountInformationService _discountService;
        #region Tìm kiếm các đợt nhập
   

        public override async Task<PagedResponseDto> Search(BaseFilter filter)
        {
            try
            {
                var statusMap = new Dictionary<string, string>
                    {
                        { "01", "khởi tạo" },
                        { "02", "Trình duyệt" },
                        { "03", "Yêu cầu chỉnh sửa" },
                        { "04", "phê duyệt" },
                        { "05", "từ chối" },
                        };
                var keyword = filter.KeyWord?.ToLower() ?? "";
                var matchedStatus= statusMap
                    .Where(x => x.Value.ToLower().Contains(keyword))
                    .Select(x => x.Key)
                    .ToList();
                var query = _dbContext.TblBuCalculateDiscount.AsQueryable();

                if (!string.IsNullOrWhiteSpace(filter.KeyWord))
                {
                    query = query.Where(x => x.Name.Contains(filter.KeyWord) || matchedStatus.Contains(x.Status) || x.Date.ToString().Contains(filter.KeyWord));
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
                var lstSigner = await _dbContext.TblMdSigner.Where(x => x.IsActive == true).ToArrayAsync();
                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.CreateDate).OrderBy(x => x.Order).ToListAsync();
                var lstMarket = await _dbContext.TblMdMarket.Where(x => x.IsActive == true).OrderBy(x => x.Code).ToListAsync();
                var lstCustomerDb = await _dbContext.TblMdCustomerDb.Where(x => x.IsActive == true).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerPt = await _dbContext.TblMdCustomerPt.Where(x => x.IsActive == true).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerFob = await _dbContext.TblMdCustomerFob.Where(x => x.IsActive == true).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerTnpp = await _dbContext.TblMdCustomerTnpp.Where(x => x.IsActive == true).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerBbdo = await _dbContext.TblMdCustomerBbdo.Where(x => x.IsActive == true).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerPts = await _dbContext.TblMdCustomerPts.Where(x => x.IsActive == true).OrderBy(x => x.Order).ToListAsync();

                return new CalculateDiscountInputModel
                {
                    Header = new TblBuCalculateDiscount
                    {
                        Id = headerId,
                        Date = DateTime.Now,
                        Hour = DateTime.Now,
                        IsActive = true,
                        CongDienPtBanLe = "",
                        CongDienSo = "",
                        QuyetDinhSo = "",
                        VanBanSo = "",
                        SignerCode = lstSigner.Where(x => x.Type == "NguoiKy" && x.IsSelect == true).FirstOrDefault()?.Code ?? "",
                        KdxdCode = lstSigner.Where(x => x.Type == "kdxd" && x.IsSelect == true).FirstOrDefault()?.Code ?? "",
                        TcktCode = lstSigner.Where(x => x.Type == "tckt" && x.IsSelect == true).FirstOrDefault()?.Code ?? "",
                        VietphuonganCode = lstSigner.Where(x => x.Type == "vietPhuongAn" && x.IsSelect == true).FirstOrDefault()?.Code ?? "",
                        Status = "01"
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
                    CustomerPts = lstCustomerPts.Select(x => new TblBuInputCustomerPts
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        CuLyBq = x.CuLyBq,
                        Cpccvc = x.Cpccvc,
                        Cvcbq = x.Cvcbq,
                        Lvnh = x.Lvnh,
                        Htcvc = x.Htcvc,
                        HttVb1370 = x.HttVb1370,
                        Ckv2 = x.Ckv2,
                        GoodsCode = x.GoodsCode,
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
                    CustomerBbdo = lstCustomerBbdo.Select(x => new TblBuInputCustomerBbdo
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = headerId,
                        Code = x.Code,
                        Name = x.Name,
                        LocalCode = x.LocalCode ?? "-",
                        DeliveryPoint = x.DeliveryPoint ?? "-",
                        DeliveryGroupCode = x.DeliveryGroupCode ?? "-",
                        GoodsCode = x.GoodsCode ?? "-",
                        MarketCode = x.MarketCode ?? "-",
                        CuLyBq = x.CuLyBq ?? 0,
                        Cpccvc = x.Cpccvc ?? 0,
                        Cvcbq = x.Cvcbq ?? 0,
                        Lvnh = x.Lvnh ?? 0,
                        Fob = x.Fob ?? 0,
                        Htcvc = x.Htcvc ?? 0,
                        HttVb1370 = x.HttVb1370 ?? 0,
                        Ckv2 = x.Ckv2 ?? 0,
                        PhuongThuc = x.PhuongThuc ?? "-",
                        Thtt = x.Thtt ?? "-",
                        Order = x.Order,
                        Adrress = x.Adrress ?? "-",
                        CkDau = x.CkDau ?? 0,
                        CkXang = x.CkXang ?? 0,
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
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
            input.Header.Date = input.Header.Date.AddHours(7);
            input.Header.Hour = (DateTime)(input.Header.Hour?.AddHours(7));
            try
            {
                _dbContext.TblBuCalculateDiscount.Add(input.Header);
                _dbContext.TblBuInputPrice.AddRange(input.InputPrice);
                _dbContext.TblBuInputMarket.AddRange(input.Market);
                _dbContext.TblBuInputCustomerDb.AddRange(input.CustomerDb);
                _dbContext.TblBuInputCustomerPt.AddRange(input.CustomerPt);
                _dbContext.TblBuInputCustomerPts.AddRange(input.CustomerPts);
                _dbContext.TblBuInputCustomerFob.AddRange(input.CustomerFob);
                _dbContext.TblBuInputCustomerTnpp.AddRange(input.CustomerTnpp);
                _dbContext.TblBuInputCustomerBbdo.AddRange(input.CustomerBbdo);
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
                    CustomerPts = await _dbContext.TblBuInputCustomerPts.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    CustomerFob = await _dbContext.TblBuInputCustomerFob.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    CustomerTnpp = await _dbContext.TblBuInputCustomerTnpp.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
                    CustomerBbdo = await _dbContext.TblBuInputCustomerBbdo.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync(),
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
                if (input.Header.Status == "01" || input.Header.Status == "03" || input.Header.Status == "04" || input.Header.Status == "08")
                {
                    _dbContext.TblBuCalculateDiscount.Update(input.Header);
                    _dbContext.TblBuInputPrice.UpdateRange(input.InputPrice);
                    _dbContext.TblBuInputMarket.UpdateRange(input.Market);
                    _dbContext.TblBuInputCustomerDb.UpdateRange(input.CustomerDb);
                    _dbContext.TblBuInputCustomerPt.UpdateRange(input.CustomerPt);
                    _dbContext.TblBuInputCustomerPts.UpdateRange(input.CustomerPts);
                    _dbContext.TblBuInputCustomerFob.UpdateRange(input.CustomerFob);
                    _dbContext.TblBuInputCustomerTnpp.UpdateRange(input.CustomerTnpp);
                    _dbContext.TblBuInputCustomerBbdo.UpdateRange(input.CustomerBbdo);
                    var h = new TblBuHistoryAction()
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = input.Header.Id,
                        Action = "Cập nhật thông tin",
                    };
                    _dbContext.TblBuHistoryAction.Add(h);

                    await _dbContext.SaveChangesAsync();
                }
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
                var lstCustomerPts = await _dbContext.TblBuInputCustomerPts.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerFob = await _dbContext.TblBuInputCustomerFob.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerTnpp = await _dbContext.TblBuInputCustomerTnpp.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var lstCustomerBbdo = await _dbContext.TblBuInputCustomerBbdo.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();

                var currentHeader = _dbContext.TblBuCalculateDiscount.Find(id);
                data.Header = currentHeader;

                var currentData = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == id).OrderBy(x => x.Order).ToListAsync();
                var previousHeader = await _dbContext.TblBuCalculateDiscount.Where(x => x.Date < currentHeader.Date && x.Status == "04").OrderByDescending(x => x.Date).FirstOrDefaultAsync();
                var previousData = new List<TblBuInputPrice>();
                data.Dlg.NameOld = "";
                if (previousHeader != null)
                {
                    previousData = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == previousHeader.Id).ToListAsync();
                    data.Dlg.NameOld = previousHeader.Name;
                    data.HeaderOld = previousHeader;
                }
                data.InputPriceOld = previousData;
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
                        Col3 = i.GblV2 - i.GblV1
                    });
                        //i.GblV2 - i.GblV1 >= 0
                        //    ? "(+)" + Math.Abs(i.GblV2 - i.GblV1)
                        //    : "(-)" + Math.Abs(i.GblV2 - i.GblV1)
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

                // Lãi gộp
                foreach (var g in lstGoods)
                {
                    var oldData = previousData.Where(x => x.GoodCode == g.Code).ToList();
                    var k = new DlgModel
                    {
                        GoodCode = g.Code,
                        LocalCode = "V1",
                        GoodName = g.Name,
                        Col1 = oldData.Sum(x => x.Vcf),
                        Col2 = oldData.Sum(x => x.ThueBvmt),
                        Col3 = oldData.Sum(x => x.L15Nbl),
                        Col4 = oldData.Sum(x => x.Vcf) * oldData.Sum(x => x.L15Nbl),
                        Col5 = (oldData.Sum(x => x.ThueBvmt) + oldData.Sum(x => x.Vcf) * oldData.Sum(x => x.L15Nbl)) * 1.1M,
                        Col6 = oldData.Sum(x => x.GblV1 + x.ChenhLech),
                        Col14 = oldData.Sum(x => x.FobV1),
                        Col10 = oldData.Sum(x => x.LaiGop) == null ? 0 : oldData.Sum(x => x.LaiGop),
                    };
                    if (k.Col6 != 0)
                    {
                        k.Col7 = k.Col6 / 1.1M - k.Col2;
                    }
                    k.Col8 = k.Col6 - k.Col5;
                    if (k.Col8 != 0)
                    {
                        k.Col9 = k.Col8 / 1.1M;
                    }
                    k.Col11 = k.Col1 * k.Col10 * 1.1M;
                    k.Col13 = k.Col11 + k.Col9;
                    k.Col12 = k.Col13 * 1.1M;
                    k.Col15 = (k.Col12 - k.Col14) * k.Col1;
                    k.Col16 = k.Col12 - k.Col14;
                    data.Dlg.Dlg6Old.Add(k);
                }
                foreach (var g in lstGoods)
                {
                    var oldData = previousData.Where(x => x.GoodCode == g.Code).ToList();
                    //var dlg1 = data.DLG.Dlg_3.Where(x => x.Code == g.Code).ToList();
                    var k = new DlgModel
                    {
                        GoodCode = g.Code,
                        LocalCode = "V2",
                        GoodName = g.Name,
                        Col1 = oldData.Sum(x => x.Vcf),
                        Col2 = oldData.Sum(x => x.ThueBvmt),
                        Col3 = oldData.Sum(x => x.L15Nbl),
                        Col4 = oldData.Sum(x => x.Vcf) * oldData.Sum(x => x.L15Nbl),
                        Col5 = (oldData.Sum(x => x.ThueBvmt) + oldData.Sum(x => x.Vcf) * oldData.Sum(x => x.L15Nbl)) * 1.1M,
                        Col6 = oldData.Sum(x => x.GblV2),
                        Col14 = oldData.Sum(x => x.FobV2),
                        Col10 = oldData.Sum(x => x.LaiGop) == null ? 0 : oldData.Sum(x => x.LaiGop),
                    };
                    if (k.Col6 != 0)
                    {
                        k.Col7 = k.Col6 / 1.1M - k.Col2;
                    }
                    k.Col8 = k.Col6 - k.Col5;
                    if (k.Col8 != 0)
                    {
                        k.Col9 = k.Col8 / 1.1M;
                    }
                    k.Col11 = k.Col1 * k.Col10 * 1.1M;
                    k.Col13 = k.Col11 + k.Col9;
                    k.Col12 = k.Col13 * 1.1M;
                    k.Col15 = (k.Col12 - k.Col14) * k.Col1;
                    k.Col16 = k.Col12 - k.Col14;
                    data.Dlg.Dlg6Old.Add(k);

                }
                foreach (var g in lstGoods)
                {
                    var dlg_6 = data.Dlg.Dlg6;
                    var dlg_6_Old = data.Dlg.Dlg6Old;
                    foreach (var n in dlg_6)
                    {
                        if (g.Code == n.GoodCode)
                        {
                            var i = new DlgModel
                            {
                                GoodCode = g.Code,
                                GoodName = g.Name,
                                LocalCode = n.LocalCode,
                                Col1 = dlg_6_Old.Where(x => x.GoodCode == g.Code).Where(x => x.LocalCode == n.LocalCode).Sum(x => x.Col12),
                                Col2 = n.Col12,
                                Col3 = n.Col12 - dlg_6_Old.Where(x => x.GoodCode == g.Code).Where(x => x.LocalCode == n.LocalCode).Sum(x => x.Col12),
                            };

                            data.Dlg.Dlg9.Add(i);
                        }
                    }
                }

                // Đề xuất mức giảm giá
                foreach (var g in lstGoods)
                {
                    var dlg_6 = data.Dlg.Dlg6;
                    var dlg_6_Old = data.Dlg.Dlg6Old;

                    foreach (var n in dlg_6)
                    {
                        if (g.Code == n.GoodCode && n.LocalCode == "V2")
                        {
                            var i = new DlgModel
                            {
                                GoodCode = g.Code,
                                GoodName = g.Name,
                                LocalCode = n.LocalCode,
                                Col1 = dlg_6_Old.Where(x => x.GoodCode == g.Code).Where(x => x.LocalCode == "V2").Sum(x => x.Col14),
                                Col2 = n.Col14,
                                Col3 = n.Col14 - dlg_6_Old.Where(x => x.GoodCode == g.Code).Where(x => x.LocalCode == "V2").Sum(x => x.Col14),
                            };
                            data.Dlg.Dlg10.Add(i);
                        }

                    }
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

                        col14 = data.Pt.Where(x => x.MarketCode == marketCode).Sum(x => Math.Round(x.Col11 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                        col16 = data.Pt.Where(x => x.MarketCode == marketCode).Sum(x => Math.Round(x.Col13 / 10, 0, MidpointRounding.AwayFromZero) * 10);

                        if (i.Code == "906962")
                        {
                            col18 = data.Pt.Where(x => x.MarketCode == "V2_02").Sum(x => Math.Round(x.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                            col20 = data.Pt.Where(x => x.MarketCode == "V2_02").Sum(x => Math.Round(x.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                        }
                        else
                        {
                            col18 = data.Pt.Where(x => x.MarketCode == marketCode).Sum(x => Math.Round(x.Col15 / 10, 0, MidpointRounding.AwayFromZero) * 10);
                            col20 = data.Pt.Where(x => x.MarketCode == marketCode).Sum(x => Math.Round(x.Col17 / 10, 0, MidpointRounding.AwayFromZero) * 10);
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

                #region BB DO
                var lstDG = await _dbContext.TblMdDeliveryGroup.OrderBy(x => x.Code).ToListAsync();
                var obb = 1;
                var _I = lstCustomerBbdo.Where(x => x.GoodsCode == "0601002").ToList();
                data.Bbdo.Add(new DataModel
                {
                    CustomerName = "I. MẶT HÀNG DẦU DO 0,05S-II",
                    IsBold = true,
                });
                foreach (var d in lstDG)
                {
                    data.Bbdo.Add(new DataModel
                    {
                        CustomerName = d.Name,
                        IsBold = true,
                    });
                    foreach (var i in _I.Where(x => x.DeliveryGroupCode == d.Code))
                    {
                        var j = new DataModel
                        {
                            Stt = obb.ToString(),
                            Id = i.Id,
                            CustomerName = i.Name,
                            DeliveryPoint = i.DeliveryPoint,
                            GoodName = "Điêzen 0,05S-II",
                            PThuc = i.PhuongThuc,
                            CustomerCode = i.Code,
                            LocalCode = i.LocalCode,
                            GoodCode = i.GoodsCode,
                            Dvt = "L",
                            TToan = i.Thtt,
                            Col1 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col12),
                            Col2 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col12) / 1.1M,
                            Col3 = i.Cpccvc + i.Cvcbq + i.Lvnh,
                            Col4 = i.Cpccvc,
                            Col5 = i.Cvcbq,
                            Col6 = i.Lvnh,
                            Col7 = i.Fob == 0 ? (decimal)((data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => (decimal?)x.Col14) ?? 0M)) + 50 : (decimal)i.Fob,
                            Col8 = (i.Fob == 0 ? (decimal)((data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => (decimal?)x.Col14) ?? 0M)) + 50 : (decimal)i.Fob) / 1.1M,
                            Col15 = i.LamTronDacBiet == true ? 1 : 0,

                        };
                        j.Col9 = j.Col7 - j.Col5 * 1.1M;
                        //Math.Round((j.Col7 - j.Col5 * 1.1M), MidpointRounding.AwayFromZero);
                        j.Col10 = j.Col9 / 1.1M;
                        j.Col12 = j.Col2 - j.Col3 - j.Col10 - j.Col6;
                        j.Col11 = j.Col12 * 1.1M;

                        var dlg6 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").FirstOrDefault();
                        if (j.LocalCode == "V1")
                        {
                            dlg6 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V1").FirstOrDefault();
                        }

                        j.Col14 = dlg6.Col6 - j.Col9;
                        j.Col14 = j.Col14 > dlg6.Col6
                            ? dlg6.Col6
                            : j.Col14;
                        //j.Col14 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col6) - j.Col9;
                        //j.Col14 = j.Col14 > data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col6) ? data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col6) : j.Col14;

                        j.Col13 = j.Col14 / 1.1M - data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col2);
                        if (j.Col15 == 1)
                        {
                            j.Col13 = Math.Round((j.Col13/ 10), MidpointRounding.AwayFromZero) * 10;
                        
                            j.Col14 = (j.Col13 + data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col2)) * 1.1M;
                        }


                        data.Bbdo.Add(j);
                        obb++;
                    }
                }

                var _II = lstCustomerBbdo.Where(x => x.GoodsCode == "0601005").ToList();
                data.Bbdo.Add(new DataModel
                {
                    CustomerName = "II. MẶT HÀNG DẦU DO 0,001S-V",
                    IsBold = true,
                });
                foreach (var i in _II)
                {
                    var j = new DataModel
                    {
                        Stt = obb.ToString(),
                        Id = i.Id,
                        CustomerName = i.Name,
                        DeliveryPoint = i.DeliveryPoint,
                        GoodName = "Điêzen 0.001S-V",
                        MarketCode = i.MarketCode,
                        LocalCode = i.LocalCode,
                        PThuc = i.PhuongThuc,
                        CustomerCode = i.Code,
                        GoodCode = i.GoodsCode,
                        Dvt = "L",
                        TToan = i.Thtt,
                        Col1 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col12),
                        Col2 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col12) / 1.1M,
                        Col3 = i.Cpccvc + i.Cvcbq + i.Lvnh,
                        Col4 = i.Cpccvc,
                        Col5 = i.Cvcbq,
                        Col6 = i.Lvnh,
                        Col7 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col14),
                        Col8 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col14) / 1.1M,
                        Col15 = i.LamTronDacBiet == true ? 1 : 0,

                    };
                    j.Col9 = j.Col7 - j.Col5 * 1.1M;
                        //Math.Round((j.Col7 - j.Col5 * 1.1M), MidpointRounding.AwayFromZero);
                    j.Col10 = j.Col9 / 1.1M;
                    j.Col12 = j.Col2 - j.Col3 - j.Col10 - j.Col6;
                    j.Col11 = j.Col12 * 1.1M;

                    var dlg6 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").FirstOrDefault();
                    if (j.LocalCode == "V1")
                    {
                        dlg6 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V1").FirstOrDefault();
                    }

                    j.Col14 = dlg6.Col6 - j.Col9;
                    j.Col14 = j.Col14 > dlg6.Col6
                        ? dlg6.Col6
                        : j.Col14;


                    j.Col13 = j.Col14 / 1.1M - data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col2);
                    
                    if (j.Col15 == 1)
                    {
                        j.Col13 = Math.Round((j.Col13 / 10), MidpointRounding.AwayFromZero) * 10;

                        j.Col14 = (j.Col13 + data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col2)) * 1.1M;
                    }
                    data.Bbdo.Add(j);

                    obb++;
                }

                var _III = lstCustomerBbdo.Where(x => x.GoodsCode != "0601005" && x.GoodsCode != "0601002").ToList();
                data.Bbdo.Add(new DataModel
                {
                    CustomerName = "III. MẶT HÀNG XĂNG",
                    IsBold = true,
                });
                foreach (var i in _III)
                {
                    var j = new DataModel
                    {
                        Stt = obb.ToString(),
                        Id = i.Id,
                        CustomerName = i.Name,
                        DeliveryPoint = i.DeliveryPoint,
                        GoodName = lstGoods.FirstOrDefault(x => x.Code == i.GoodsCode)?.Name,
                        PThuc = i.PhuongThuc,
                        CustomerCode = i.Code,
                        LocalCode = i.LocalCode,
                        GoodCode = i.GoodsCode,
                        Dvt = "L",
                        TToan = i.Thtt,
                        Col1 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col12),
                        Col2 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col12) / 1.1M,
                        Col3 = i.Cpccvc + i.Cvcbq + i.Lvnh,
                        Col4 = i.Cpccvc,
                        Col5 = i.Cvcbq,
                        Col6 = i.Lvnh,
                        Col7 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col14),
                        Col8 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col14) / 1.1M,
                        Col15 = i.LamTronDacBiet == true ? 1 : 0,
                    };
                    j.Col9 = j.Col7 - j.Col5 * 1.1M;
                    j.Col10 = j.Col9 / 1.1M;
                    j.Col12 = j.Col2 - j.Col3 - j.Col10 - j.Col6;
                    j.Col11 = j.Col12 * 1.1M;


                    var dlg6 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").FirstOrDefault();
                    if ( j.LocalCode == "V1")
                    {
                        dlg6 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V1").FirstOrDefault();
                    }

                    j.Col14 = dlg6.Col6 - j.Col9;
                    j.Col14 = j.Col14 > dlg6.Col6
                        ? dlg6.Col6
                        : j.Col14;

                    //j.Col14 = data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col6) - j.Col9;
                    //j.Col14 = j.Col14 > data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col6) 
                    //    ? data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col6) 
                    //    : j.Col14;

                    j.Col13 = j.Col14 / 1.1M - data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col2);
                    if (j.Col15 == 1)
                    {
                        j.Col13 = Math.Round((j.Col13 / 10), MidpointRounding.AwayFromZero) * 10;

                        j.Col14 = (j.Col13 + data.Dlg.Dlg6.Where(x => x.GoodCode == i.GoodsCode && x.LocalCode == "V2").Sum(x => x.Col2)) * 1.1M;
                    }

                    data.Bbdo.Add(j);
                    obb++;
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
                    foreach (var i in market)
                    {
                        data.Pl1.Add(new DataModel
                        {
                            Stt = _pl1.ToString(),
                            MarketCode = i.Code,
                            MarketName = i.FullName,

                            //Col1 = Math.Round(data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => x.Col11)),
                            //Col2 = Math.Round(data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => x.Col13)),
                            //Col3 = Math.Round(data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => x.Col15)),
                            //Col4 = Math.Round(data.Pt.Where(x => x.MarketCode == i.Code).Sum(x => x.Col17)),
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

                #region PL2
                var local2 = lstCustomerDb.Where(x => !string.IsNullOrEmpty(x.Local2)).Select(x => x.Local2).Distinct().ToList();
                var _pl2_lm = 1;
                foreach (var l in local2)
                {
                    data.Pl2.Add(new DataModel
                    {
                        Stt = IntToRoman(_pl2_lm),
                        CustomerName = l,
                        IsBold = true,
                    });
                    var customer = lstCustomerDb.Where(x => x.Local2 == l).ToList();
                    var _pl2 = 1;
                    foreach (var i in customer)
                    {
                        data.Pl2.Add(new DataModel
                        {
                            Stt = _pl2.ToString(),
                            CustomerCode = i.Code,
                            CustomerName = i.Name,
                            MarketCode = i.MarketCode,
                            LocalCode = i.LocalCode,
                            Col1 = Math.Round(data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col14)),
                            Col2 = Math.Round(data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col16)),
                            Col3 = Math.Round(data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col18)),
                            Col4 = Math.Round(data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => x.Col20)),
                            //Col1 = data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => Math.Round(x.Col14 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                            //Col2 = data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => Math.Round(x.Col16 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                            //Col3 = data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => Math.Round(x.Col18 / 10, 0, MidpointRounding.AwayFromZero) * 10),
                            //Col4 = data.Db.Where(x => x.CustomerCode == i.Code).Sum(x => Math.Round(x.Col20 / 10, 0, MidpointRounding.AwayFromZero) * 10),

                        });
                        _pl2++;
                    }
                    _pl2_lm++;
                }
                #endregion

                #region PL3
                var pl3 = 1;
                foreach (var i in lstCustomerFob)
                {
                    data.Pl3.Add(new DataModel
                    {
                        Stt = pl3.ToString(),
                        CustomerCode = i.Code,
                        LocalCode = i.LocalCode,
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
                        LocalCode = i.LocalCode,
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
                            Col14 = currentHeader.Date.ToString("HH:mm"),
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
                            Col14 = currentHeader.Date.ToString("HH:mm"),
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
                            Col14 = currentHeader.Date.ToString("HH:mm"),
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
                            Col14 = currentHeader.Date.ToString("HH:mm"),
                        });
                        o++;
                    }
                }
                #endregion

                #region PTS

                var Pts = 1;
                foreach (var i in lstCustomerPts)
                {
                    data.Pts.Add(new DataModel
                    {
                        Stt = Pts.ToString(),
                        CustomerName = i.Name,
                        DeliveryPoint = i.Adrress,
                        Col1 = i.CuLyBq,
                        Col2 = i.Cvcbq,
                        Col3 = data.Dlg.Dlg8.Where(x => x.GoodCode == i.GoodsCode).Sum(x => x.Col4),
                        PThuc = i.PhuongThuc,
                        CustomerCode = i.Code,
                        GoodCode = i.GoodsCode,
                        MarketName = currentHeader.Date.ToString("dd.MM.yyyy"),
                        TToan = i.Thtt,
                    });
                    Pts++;
                }

                #endregion

                #region VK11-BB
                var bb = 1;
                foreach (var i in lstCustomerBbdo)
                {
                    data.Vk11Bb.Add(new VK11Model
                    {
                        Stt = bb.ToString(),
                        CustomerName = i.Name,
                        Address = i.DeliveryPoint,
                        GoodsName = lstGoods.FirstOrDefault(x => x.Code == i.GoodsCode)?.Name ?? "",
                        Col1 = i.CuLyBq,
                        Col2 = i.Cvcbq,
                        Col3 = i.PhuongThuc,
                        Col4 = i.Code,
                        Col5 = i.GoodsCode,
                        Col7 = i.Thtt,
                        Col8 = Math.Round(data.Bbdo.Where(x => x.Id == i.Id).Sum(x => x.Col13)),
                        Col13 = currentHeader.Date.ToString("dd.MM.yyyy"),
                        Col14 = currentHeader.Date.ToString("HH:mm"),
                    });
                    bb++;
                }
                #endregion

                #region Tổng hợp
                data.Summary.Add(new VK11Model
                {
                    Stt = "A",
                    CustomerName = "TNQTM",
                    IsBold = true,
                });
                data.Summary.AddRange(data.Vk11Pt);

                data.Summary.Add(new VK11Model
                {
                    Stt = "B",
                    CustomerName = "KHÁCH ĐẶC BIỆT",
                    IsBold = true,
                });
                data.Summary.AddRange(data.Vk11Db);

                data.Summary.Add(new VK11Model
                {
                    Stt = "C",
                    CustomerName = "BÁN FOB",
                    IsBold = true,
                });
                data.Summary.AddRange(data.Vk11Fob);

                data.Summary.Add(new VK11Model
                {
                    Stt = "D",
                    CustomerName = "TNPP",
                    IsBold = true,
                });
                data.Summary.AddRange(data.Vk11Tnpp);

                data.Summary.Add(new VK11Model
                {
                    Stt = "E",
                    CustomerName = "BÁN BUÔN",
                    IsBold = true,
                });

                foreach (var i in data.Vk11Bb)
                {
                    data.Summary.Add(new VK11Model
                    {
                        IsBold = i.IsBold,
                        Stt = i.Stt,
                        CustomerName = i.CustomerName,
                        Address = i.Address,
                        MarketName = i.GoodsName,
                        Col1 = i.Col1,
                        Col2 = i.Col2,
                        Col3 = i.Col3,
                        Col4 = i.Col4,
                        Col5 = i.Col5,
                        Col7 = i.Col7,
                        Col8 = i.Col8,
                        Col13 = i.Col13,
                        Col14 = i.Col14,
                    });
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
                i.Col3 = Math.Round(i.Col3);
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
            #region pl1

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

            #region BB DO
            foreach (var i in data.Bbdo)
            {
                if (i.Col15 == 1)
                {
                    i.Col1 = Math.Round(i.Col1);
                    i.Col2 = Math.Round(i.Col2);
                    i.Col8 = Math.Round(i.Col8);
                    i.Col9 = Math.Round(i.Col9);
                    i.Col10 = Math.Round(i.Col10);
                    i.Col11 = Math.Round(i.Col11);
                    i.Col12 = Math.Round(i.Col12);
                    i.Col13 = Math.Round(i.Col13);
                    i.Col14 = Math.Round(i.Col14);
                }

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
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),(100, "C"), (90, "XC"),
                (50, "L"), (40, "XL"),(10, "X"), (9, "IX"), (5, "V"), (4, "IV"),(1, "I")
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

        #region Xuất Excel (Tất cả sheet)
        public async Task<string> ExportExcel(string headerId)
        {
            try
            {
                var header = _dbContext.TblBuCalculateDiscount.Find(headerId);
                var nguoiKy = _dbContext.TblMdSigner.Where(x => x.Code == header.SignerCode).FirstOrDefault();
                var lapBieu = _dbContext.TblMdSigner.Where(x => x.Code == header.VietphuonganCode).FirstOrDefault();
                var kdxd = _dbContext.TblMdSigner.Where(x => x.Code == header.KdxdCode).FirstOrDefault();
                var tckt = _dbContext.TblMdSigner.Where(x => x.Code == header.TcktCode).FirstOrDefault();
                var data = await this.CalculateDiscountOutput(headerId);
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template", "CoSoTinhMucGiamGia.xlsx");
                var muaMien = "";
                var tuThang = "";
               
                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException("Không tìm thấy template Excel.", templatePath);
                }

                using var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                IWorkbook workbook = new XSSFWorkbook(file);

                var styles = new
                {
                    FreeText = ExcelNPOIExtention.SetCellFreeStyle(workbook, true, HorizontalAlignment.Center, true, 16),
                    Text = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Left, true),
                    TextRight = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Right, true),
                    TextBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Left, true),
                    TextCenter = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Center, false),
                    TextCenterBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false),
                    Number = ExcelNPOIExtention.SetCellStyleNumber(workbook, false, HorizontalAlignment.Right, true),
                    NumberBold = ExcelNPOIExtention.SetCellStyleNumber(workbook, true, HorizontalAlignment.Right, true),
                };
                if(header.Date.Month > 4 && header.Date.Month < 11)
                {
                    muaMien = "Hè Thu";
                    tuThang = "5 - 10";
                }
                else
                {
                    muaMien = "Đông xuân";
                    tuThang = "11 - 4";

                }
               //if(accountGroup!= "G_NV_K") { 
                
             
                #region Dữ liệu gốc
                var sheetDlg = workbook.GetSheetAt(0);

                #region thông báo
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(1) ?? sheetDlg.CreateRow(1), 6, $"THÔNG BÁO GIÁ BÁN LẺ TỪ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.FreeText);
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(2) ?? sheetDlg.CreateRow(2), 18, $"(Từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")})", ExcelNPOIExtention.SetCellFreeStyle(workbook, true, HorizontalAlignment.Center, false, 16));
                int rowIndexDl1 = 5;
                foreach (var i in data.Dlg.Dlg1)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl1) ?? sheetDlg.CreateRow(rowIndexDl1);
                    ExcelNPOIExtention.SetCellValueNumber(row, 1, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 2, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 4, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));
                    rowIndexDl1++;
                }
                int rowIndexDl2 = 5;
                foreach (var i in data.Dlg.Dlg2)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl2) ?? sheetDlg.CreateRow(rowIndexDl2);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.GoodName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col2, number);
                    rowIndexDl2++;
                }
                int rowIndexDl3 = 5;
                foreach (var i in data.Dlg.Dlg3)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl3) ?? sheetDlg.CreateRow(rowIndexDl3);
                    if (i.GoodCode == null)
                    {
                        ExcelNPOIExtention.SetCellValueText(row, 18, i.GoodName, styles.TextBold);
                    }
                    else
                    {
                        ExcelNPOIExtention.SetCellValueText(row, 18, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 20, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueText(row, 24, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                        //ExcelNPOIExtention.SetCellValueNumber(row, 24, i.Col3, number);
                    }
                    rowIndexDl3++;
                }
                #endregion

                #region BIỂU TỔNG HỢP CÁC CHỈ TIÊU DẦU SÁNG (PT bán lẻ - V2)

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(34) ?? sheetDlg.CreateRow(34), 0, $"Tính từ: {header.Hour?.ToString("HH:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số {header.CongDienPtBanLe?.ToString() ?? ""} ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                int rowIndexDl4 = 40;

                #region xuất người ký
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(53), 0, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(53), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(53), 6, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 0, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 6, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 0, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 6, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 0, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 6, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                if (header.SignerCode == "TongGiamDoc")
                {
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(46) ?? sheetDlg.CreateRow(46), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(47) ?? sheetDlg.CreateRow(47), 9, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(53), 9, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(92) ?? sheetDlg.CreateRow(92), 11, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(93) ?? sheetDlg.CreateRow(93), 11, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 11, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(128) ?? sheetDlg.CreateRow(128), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(129) ?? sheetDlg.CreateRow(129), 9, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 9, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(164) ?? sheetDlg.CreateRow(164), 9, $"Vinh, Ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(165) ?? sheetDlg.CreateRow(165), 9, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 9, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                }
                else
                {
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(46) ?? sheetDlg.CreateRow(46), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(47) ?? sheetDlg.CreateRow(47), 9, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(48) ?? sheetDlg.CreateRow(48), 9, $"{nguoiKy?.Position ?? ""}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(53), 9, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(92) ?? sheetDlg.CreateRow(92), 11, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(93) ?? sheetDlg.CreateRow(93), 11, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(94) ?? sheetDlg.CreateRow(94), 11, $"{nguoiKy?.Position ?? ""}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 11, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(128) ?? sheetDlg.CreateRow(128), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(129) ?? sheetDlg.CreateRow(129), 9, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(130) ?? sheetDlg.CreateRow(130), 9, $"{nguoiKy?.Position ?? ""}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 9, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(164) ?? sheetDlg.CreateRow(164), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(165) ?? sheetDlg.CreateRow(165), 9, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(166) ?? sheetDlg.CreateRow(166), 9, $"{nguoiKy?.Position ?? ""}", styles.TextCenter);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 9, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                }
                #endregion

                foreach (var i in data.Dlg.Dlg4)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl4) ?? sheetDlg.CreateRow(rowIndexDl4);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 2, i.Col1, styles.TextRight);
                    ExcelNPOIExtention.SetCellValueNumber(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col4, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col5, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col6, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col7, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col10, number);
                    rowIndexDl4++;
                }
                #endregion

                #region THAY ĐỔI GIÁ GIAO PT BÁN LẺ
                int rowIndexDl5 = 39;
                foreach (var i in data.Dlg.Dlg3)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl5) ?? sheetDlg.CreateRow(rowIndexDl5);
                    if (i.GoodCode == null)
                    {
                        ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, styles.TextBold);
                    }
                    else
                    {
                        ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                        //ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                        ExcelNPOIExtention.SetCellValueText(row, 24, i.Col3 == 0 ? "(Không thay đổi)" : "(Thay đổi)", text);
                    }
                    rowIndexDl5++;
                }
                #endregion

                #region BIỂU TỔNG HỢP CÁC CHỈ TIÊU DẦU SÁNG (ngoài bán lẻ)

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(74) ?? sheetDlg.CreateRow(74), 0, $"Tính từ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số  {header.CongDienSo.ToString()}   ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                int rowIndexDl6 = 81;
               
                foreach (var i in data.Dlg.Dlg6)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl6) ?? sheetDlg.CreateRow(rowIndexDl6);
                    if (i.GoodCode == null)
                    {
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, styles.TextBold);
                    }
                    else
                    {
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);

                        ExcelNPOIExtention.SetCellValueText(row, 2, i.Col1, styles.TextRight);
                        ExcelNPOIExtention.SetCellValueNumber(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col4, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col5, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col6, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col7, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col8, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col9, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col10, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col11, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col13, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 16, i.Col15, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col16, number);
                    }
                    rowIndexDl6++;
                }
                #endregion

                #region BIỂU TÍNH GIÁ XUẤT NỘI DỤNG

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(115) ?? sheetDlg.CreateRow(115), 0, $"Tính từ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số  {header.CongDienSo.ToString()}  ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                int rowIndexDl7 = 121;
                foreach (var i in data.Dlg.Dlg7)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl7) ?? sheetDlg.CreateRow(rowIndexDl7);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);

                    ExcelNPOIExtention.SetCellValueText(row, 3, i.Col1, styles.TextRight);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col3, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col4, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col5, number);
                    
                    rowIndexDl7++;
                }
                #endregion

                #region BIỂU TÍNH GIÁ BÁN CHO CÔNG TY CP VẬN TẢI VÀ DỊCH VỤ PETROLIMEX NGHỆ TĨNH

                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(153) ?? sheetDlg.CreateRow(153), 0, $"Tính từ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số {header.CongDienSo.ToString()} ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                int rowIndexDl8 = 159;
                //if (header.SignerCode == "TongGiamDoc")
                foreach (var i in data.Dlg.Dlg8)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl8) ?? sheetDlg.CreateRow(rowIndexDl8);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);

                    ExcelNPOIExtention.SetCellValueText(row, 3, i.Col1, styles.TextRight);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col3, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col4, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col5, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col6, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col7, number);

                    rowIndexDl8++;
                }
                #endregion

                #region SO SÁNH

                #region thay đổi) giá giao phương thức bán lẻ
                int rowIndexDl13 = 53;
                var check = 159;
                foreach (var i in data.Dlg.Dlg5)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl13) ?? sheetDlg.CreateRow(rowIndexDl13);
                    ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                    //ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);   
                    ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                    ExcelNPOIExtention.SetCellValueText(sheetDlg.GetRow(check) ?? sheetDlg.CreateRow(check), 12, i.Col3 == 0 ? "(Không thay đổi)" : "(Thay đổi)", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Center, true, 13));
                    check++;
                    rowIndexDl13++;
                }
                #endregion

                ExcelNPOIExtention.SetCellValueText(sheetDlg.GetRow(73) ?? sheetDlg.CreateRow(73), 20, $"{header.Date.ToString("dd.MM.yyyy")}", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Center, true, 13));
                ExcelNPOIExtention.SetCellValueText(sheetDlg.GetRow(73) ?? sheetDlg.CreateRow(73), 21, $"{header.Date.ToString("hh:mm")}", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Center, true, 13));

                #region thay đổi) lãi gộp
                int rowIndexDl9_1 = 81;
                int rowIndexDl9_2 = 86;
                ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(76) ?? sheetDlg.CreateRow(76), 20, $"Lãi gộp từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} và tính theo VCF {muaMien} từ tháng {tuThang} hàng năm", styles.TextCenter);

                foreach (var i in data.Dlg.Dlg9)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    if ( i.LocalCode == "V1" )
                    {
                        var row = sheetDlg.GetRow(rowIndexDl9_1) ?? sheetDlg.CreateRow(rowIndexDl9_1);
                        ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                    //ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));
                        ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);

                        rowIndexDl9_1++;
                    }
                    else
                    {
                        var row = sheetDlg.GetRow(rowIndexDl9_2) ?? sheetDlg.CreateRow(rowIndexDl9_2);

                        ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);
                        //ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                        rowIndexDl9_2++;
                    }

                }
                #endregion

                #region thay đổi chiết khấu
                int rowIndexDl10 = 95;
                foreach (var i in data.Dlg.Dlg10)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDlg.GetRow(rowIndexDl10) ?? sheetDlg.CreateRow(rowIndexDl10);
                    ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                    //ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number); ExcelNPOIExtention.SetCellValueText(row, 4, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));


                    rowIndexDl10++;
                }
                    #endregion

                    #endregion

                    #endregion
                //}

                #region PT
                var sheetPt = workbook.GetSheetAt(1);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(1) ?? sheetPt.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPt = 7;
                foreach (var i in data.Pt)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPt.GetRow(rowIndexPt) ?? sheetPt.CreateRow(rowIndexPt);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text); 
                    ExcelNPOIExtention.SetCellValue(row, 1, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col4, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col5, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col6, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col7, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col9, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col10, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col11, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col13, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 16, i.Col15, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col16, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 18, i.Col17, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 19, i.Col18, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 20, i.Col19, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col20, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col21, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col22, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 24, i.Col23, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 25, i.Col24, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 26, i.Col25, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 27, i.Col26, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 28, i.Col27, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 29, i.Col28, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 30, i.Col29, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 31, i.Col30, number);
                    rowIndexPt++;
                }   
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 15, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 5) ?? sheetPt.CreateRow(rowIndexPt + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 5) ?? sheetPt.CreateRow(rowIndexPt + 5), 5, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 5) ?? sheetPt.CreateRow(rowIndexPt + 5), 10, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 5) ?? sheetPt.CreateRow(rowIndexPt + 5), 15, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));


                #endregion

                #region ĐB
                var sheetDb = workbook.GetSheetAt(2);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(1) ?? sheetDb.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexDb = 7;
                foreach (var i in data.Db)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDb.GetRow(rowIndexDb) ?? sheetDb.CreateRow(rowIndexDb);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValue(row, 2, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 6, i.Col4, number);
                    ExcelNPOIExtention.SetCellValue(row, 7, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col13, number);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col15, number);
                    ExcelNPOIExtention.SetCellValue(row, 18, i.Col16, number);
                    ExcelNPOIExtention.SetCellValue(row, 19, i.Col17, number);
                    ExcelNPOIExtention.SetCellValue(row, 20, i.Col18, number);
                    ExcelNPOIExtention.SetCellValue(row, 21, i.Col19, number);
                    ExcelNPOIExtention.SetCellValue(row, 22, i.Col20, number);
                    ExcelNPOIExtention.SetCellValue(row, 23, i.Col21, number);
                    ExcelNPOIExtention.SetCellValue(row, 24, i.Col22, number);
                    ExcelNPOIExtention.SetCellValue(row, 25, i.Col23, number);
                    ExcelNPOIExtention.SetCellValue(row, 26, i.Col24, number);
                    ExcelNPOIExtention.SetCellValue(row, 27, i.Col25, number);
                    ExcelNPOIExtention.SetCellValue(row, 28, i.Col26, number);
                    ExcelNPOIExtention.SetCellValue(row, 29, i.Col27, number);
                    ExcelNPOIExtention.SetCellValue(row, 30, i.Col28, number);
                    ExcelNPOIExtention.SetCellValue(row, 31, i.Col29, number);
                    ExcelNPOIExtention.SetCellValue(row, 32, i.Col30, number);
                    ExcelNPOIExtention.SetCellValue(row, 33, i.Col31, number);
                    ExcelNPOIExtention.SetCellValue(row, 34, i.Col32, number);
                    ExcelNPOIExtention.SetCellValue(row, 35, i.Col33, number);
                    ExcelNPOIExtention.SetCellValue(row, 36, i.Col34, number);
                    rowIndexDb++;
                }
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 15, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 5) ?? sheetDb.CreateRow(rowIndexDb + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 5) ?? sheetDb.CreateRow(rowIndexDb + 5), 5, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 5) ?? sheetDb.CreateRow(rowIndexDb + 5), 10, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 5) ?? sheetDb.CreateRow(rowIndexDb + 5), 15, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));


                #endregion

                #region FOB
                var sheetFob = workbook.GetSheetAt(3);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(1) ?? sheetFob.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexFob = 7;
                foreach (var i in data.Fob)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetFob.GetRow(rowIndexFob) ?? sheetFob.CreateRow(rowIndexFob);

                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);
                    ExcelNPOIExtention.SetCellValue(row, 6, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 7, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col12, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col13, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col14, number);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col15, number);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col16, number);
                    ExcelNPOIExtention.SetCellValue(row, 18, i.Col17, number);
                    ExcelNPOIExtention.SetCellValue(row, 19, i.Col18, number);
                    ExcelNPOIExtention.SetCellValue(row, 20, i.Col19, number);
                    ExcelNPOIExtention.SetCellValue(row, 21, i.Col20, number);
                    ExcelNPOIExtention.SetCellValue(row, 22, i.Col21, number);
                    ExcelNPOIExtention.SetCellValue(row, 23, i.Col22, number);
                    ExcelNPOIExtention.SetCellValue(row, 24, i.Col23, number);
                    ExcelNPOIExtention.SetCellValue(row, 25, i.Col24, number);
                    ExcelNPOIExtention.SetCellValue(row, 26, i.Col25, number);
                    ExcelNPOIExtention.SetCellValue(row, 27, i.Col26, number);
                    ExcelNPOIExtention.SetCellValue(row, 28, i.Col27, number);
                    ExcelNPOIExtention.SetCellValue(row, 29, i.Col28, number);
                    ExcelNPOIExtention.SetCellValue(row, 30, i.Col29, number);
                    ExcelNPOIExtention.SetCellValue(row, 31, i.Col30, number);
                    ExcelNPOIExtention.SetCellValue(row, 32, i.Col31, number);
                    ExcelNPOIExtention.SetCellValue(row, 33, i.Col32, number);

                    rowIndexFob++;
                }
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 15, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 5) ?? sheetFob.CreateRow(rowIndexFob + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 5) ?? sheetFob.CreateRow(rowIndexFob + 5), 5, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 5) ?? sheetFob.CreateRow(rowIndexFob + 5), 10, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 5) ?? sheetFob.CreateRow(rowIndexFob + 5), 15, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region PT 09
                var sheetPt09 = workbook.GetSheetAt(4);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(1) ?? sheetPt09.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPt09 = 7;
                foreach (var i in data.Pt09)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPt09.GetRow(rowIndexPt09) ?? sheetPt09.CreateRow(rowIndexPt09);

                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);
                    ExcelNPOIExtention.SetCellValue(row, 6, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 7, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col12, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col13, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col14, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col15, number);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col16, number);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col17, number);
                    ExcelNPOIExtention.SetCellValue(row, 18, i.Col18, number);
                    ExcelNPOIExtention.SetCellValue(row, 19, i.Col19, number);
                    ExcelNPOIExtention.SetCellValue(row, 20, i.Col20, number);
                    ExcelNPOIExtention.SetCellValue(row, 21, i.Col21, number);
                    ExcelNPOIExtention.SetCellValue(row, 22, i.Col22, number);
                    ExcelNPOIExtention.SetCellValue(row, 23, i.Col23, number);
                    ExcelNPOIExtention.SetCellValue(row, 24, i.Col24, number);
                    ExcelNPOIExtention.SetCellValue(row, 25, i.Col25, number);
                    ExcelNPOIExtention.SetCellValue(row, 26, i.Col26, number);
                    ExcelNPOIExtention.SetCellValue(row, 27, i.Col27, number);
                    ExcelNPOIExtention.SetCellValue(row, 28, i.Col28, number);
                    ExcelNPOIExtention.SetCellValue(row, 29, i.Col29, number);
                    ExcelNPOIExtention.SetCellValue(row, 30, i.Col30, number);
                    ExcelNPOIExtention.SetCellValue(row, 31, i.Col31, number);
                    ExcelNPOIExtention.SetCellValue(row, 32, i.Col32, number);
                    //ExcelNPOIExtention.SetCellValue(row, 33, i.Col33, number);

                    rowIndexPt09++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 15, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 5) ?? sheetPt09.CreateRow(rowIndexPt09 + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 5) ?? sheetPt09.CreateRow(rowIndexPt09 + 5), 5, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 5) ?? sheetPt09.CreateRow(rowIndexPt09 + 5), 10, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 5) ?? sheetPt09.CreateRow(rowIndexPt09 + 5), 15, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region BB DO
                var sheetBbDo = workbook.GetSheetAt(5);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(3) ?? sheetBbDo.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexBbDo = 9;
                foreach (var i in data.Bbdo)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetBbDo.GetRow(rowIndexBbDo) ?? sheetBbDo.CreateRow(rowIndexBbDo);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.DeliveryPoint, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 4, i.PThuc, text);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.CustomerCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.GoodCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Dvt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.TToan, text);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col4, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 18, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 19, i.Col11 < 0 ? $"({i.Col11 * -1})" : i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 20, i.Col12 < 0 ? $"({i.Col12 * -1})" : i.Col12, number);
                    ExcelNPOIExtention.SetCellValue(row, 21, i.Col13, number);
                    ExcelNPOIExtention.SetCellValue(row, 22, i.Col14, number);
                    rowIndexBbDo++;
                }
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 15, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 5) ?? sheetBbDo.CreateRow(rowIndexBbDo + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 5) ?? sheetBbDo.CreateRow(rowIndexBbDo + 5), 5, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 5) ?? sheetBbDo.CreateRow(rowIndexBbDo + 5), 10, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 5) ?? sheetBbDo.CreateRow(rowIndexBbDo + 5), 15, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region BB FO

                #endregion

                #region PL1
                var sheetPl1 = workbook.GetSheetAt(7);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(2) ?? sheetPl1.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl1 = 8;
                foreach (var i in data.Pl1)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl1.GetRow(rowIndexPl1) ?? sheetPl1.CreateRow(rowIndexPl1);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.MarketName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl1++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 2, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 4, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 7, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 5) ?? sheetPl1.CreateRow(rowIndexPl1 + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 5) ?? sheetPl1.CreateRow(rowIndexPl1 + 5), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 5) ?? sheetPl1.CreateRow(rowIndexPl1 + 5), 4, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 5) ?? sheetPl1.CreateRow(rowIndexPl1 + 5), 7, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region PL2

                var sheetPl2 = workbook.GetSheetAt(8);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(2) ?? sheetPl2.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl2 = 7;
                foreach (var i in data.Pl2)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl2.GetRow(rowIndexPl2) ?? sheetPl2.CreateRow(rowIndexPl2);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl2++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 2, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 4, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 7, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl2 + 5) ?? sheetPl1.CreateRow(rowIndexPl2 + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl2 + 5) ?? sheetPl1.CreateRow(rowIndexPl2 + 5), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl2 + 5) ?? sheetPl1.CreateRow(rowIndexPl2 + 5), 4, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl2 + 5) ?? sheetPl1.CreateRow(rowIndexPl2 + 5), 7, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region PL3

                var sheetPl3 = workbook.GetSheetAt(9);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(2) ?? sheetPl3.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl3 = 7;
                foreach (var i in data.Pl3)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl3.GetRow(rowIndexPl3) ?? sheetPl3.CreateRow(rowIndexPl3);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl3++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 2, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 4, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 7, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 5) ?? sheetPl3.CreateRow(rowIndexPl3 + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 5) ?? sheetPl3.CreateRow(rowIndexPl3 + 5), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 5) ?? sheetPl3.CreateRow(rowIndexPl3 + 5), 4, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 5) ?? sheetPl3.CreateRow(rowIndexPl3 + 5), 7, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region PL4

                var sheetPl4 = workbook.GetSheetAt(10);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(3) ?? sheetPl4.CreateRow(3), 2, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl4 = 8;
                foreach (var i in data.Pl4)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl4.GetRow(rowIndexPl4) ?? sheetPl4.CreateRow(rowIndexPl4);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl4++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 2, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 4, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 7, "DUYỆT", styles.TextCenterBold);

                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 5) ?? sheetPl4.CreateRow(rowIndexPl4 + 5), 1, $"{lapBieu?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 5) ?? sheetPl4.CreateRow(rowIndexPl4 + 5), 2, $"{kdxd?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 5) ?? sheetPl4.CreateRow(rowIndexPl4 + 5), 4, $"{tckt?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 5) ?? sheetPl4.CreateRow(rowIndexPl4 + 5), 7, $"{nguoiKy?.Name ?? ""}", ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false));

                #endregion

                #region VK11-PT

                var sheetVk11Pt = workbook.GetSheetAt(11);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Pt.GetRow(3) ?? sheetVk11Pt.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Pt = 2;
                foreach (var i in data.Vk11Pt)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Pt.GetRow(rowIndexVk11Pt) ?? sheetVk11Pt.CreateRow(rowIndexVk11Pt);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueText(row, 9, i.Col6, text);
                    ExcelNPOIExtention.SetCellValueText(row, 10, i.Col7, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueText(row, 12, i.Col9, text);
                    ExcelNPOIExtention.SetCellValueText(row, 13, i.Col10, text);
                    ExcelNPOIExtention.SetCellValueText(row, 14, i.Col11, text);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.Col12, text);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValueText(row, 17, i.Col14, text);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexVk11Pt++;
                }
                #endregion

                #region VK11-ĐB

                var sheetVk11Db = workbook.GetSheetAt(12);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Db.GetRow(3) ?? sheetVk11Db.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Db = 2;
                foreach (var i in data.Vk11Db)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Db.GetRow(rowIndexVk11Db) ?? sheetVk11Db.CreateRow(rowIndexVk11Db);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValue(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 4, i.Col1, text);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.Col2, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexVk11Db++;
                }
                #endregion

                #region VK11-FOB

                var sheetVk11Fob = workbook.GetSheetAt(13);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Fob.GetRow(3) ?? sheetVk11Fob.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Fob = 3;
                foreach (var i in data.Vk11Fob)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Fob.GetRow(rowIndexVk11Fob) ?? sheetVk11Fob.CreateRow(rowIndexVk11Fob);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 3, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, "", number);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueText(row, 9, i.Col6, text);
                    ExcelNPOIExtention.SetCellValueText(row, 10, i.Col7, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col10, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col11, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexVk11Fob++;
                }
                #endregion

                #region VK11-TNPP

                var sheetVk11Tnpp = workbook.GetSheetAt(14);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Tnpp.GetRow(3) ?? sheetVk11Tnpp.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Tnpp = 3;
                foreach (var i in data.Vk11Tnpp)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Tnpp.GetRow(rowIndexVk11Tnpp) ?? sheetVk11Tnpp.CreateRow(rowIndexVk11Tnpp);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col6, text);
                    ExcelNPOIExtention.SetCellValueText(row, 9, i.Col7, text);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.Col13, text);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 17, i.Col15, text);
                    rowIndexVk11Tnpp++;
                }
                #endregion

                #region VK11-BB

                var sheetVk11Bb = workbook.GetSheetAt(15);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Bb.GetRow(3) ?? sheetVk11Bb.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Bb = 2;
                foreach (var i in data.Vk11Bb)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Bb.GetRow(rowIndexVk11Bb) ?? sheetVk11Bb.CreateRow(rowIndexVk11Bb);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodsName, text);
                    ExcelNPOIExtention.SetCellValue(row, 4, "", number);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col5, text);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.Col13, text);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 17, i.Col15, text);
                    rowIndexVk11Bb++;
                }
                #endregion

                #region PTS
                var sheetPts = workbook.GetSheetAt(16);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Pt.GetRow(3) ?? sheetVk11Pt.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPts = 4;
                foreach (var i in data.Pts)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPts.GetRow(rowIndexPts) ?? sheetPts.CreateRow(rowIndexPts);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.PThuc, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.CustomerCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 4, "L15", text);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.TToan, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 6, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 7, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col3, number);
                    ExcelNPOIExtention.SetCellValueText(row, 10, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 11, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 12, "L15", text);
                    ExcelNPOIExtention.SetCellValueText(row, 13, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 14, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 16, "31.12.9999", text);
                    ExcelNPOIExtention.SetCellValueText(row, 17, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 18, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 19, "PTS", text);

                    rowIndexPts++;
                }
                #endregion

                #region Tổng hợp

                var sheetTh = workbook.GetSheetAt(17);
                //ExcelNPOIExtention.SetCellValue(sheetTh.GetRow(3) ?? sheetTh.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexTh = 2;
                foreach (var i in data.Summary)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetTh.GetRow(rowIndexTh) ?? sheetTh.CreateRow(rowIndexTh);

                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col6, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col7, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col10, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col11, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexTh++;
                }
                #endregion




                var folderPath = Path.Combine($"Uploads/Excel/{DateTime.Now.ToString("yyyy/MM/dd")}");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);
                var fileName = $"CSTMGG_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                var outputPath = Path.Combine(Directory.GetCurrentDirectory(), folderPath, fileName);

                using var outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFile);

                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = fileName,
                    Type = "xlsx",
                    Path = $"{folderPath}/{fileName}",
                });
                await _dbContext.SaveChangesAsync();

                return $"{folderPath}/{fileName}";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi xuất file Excel: {ex.Message}\n{ex.StackTrace}");
                return string.Empty;
            }
        }
        #endregion
        
        #region xuất excel TPKD
        public async Task<string> ExportExcelTPKD(string headerId)
        {
            
                try
                {
                    var header = _dbContext.TblBuCalculateDiscount.Find(headerId);
                    var nguoiKy = _dbContext.TblMdSigner.Where(x => x.Code == header.SignerCode).FirstOrDefault();
                    var data = await this.CalculateDiscountOutput(headerId);
                    var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template", "CoSoTinhMucGiamGiaTPKD.xlsx");
                    var muaMien = "";
                    var tuThang = "";

                    if (!File.Exists(templatePath))
                    {
                        throw new FileNotFoundException("Không tìm thấy template Excel.", templatePath);
                    }

                    using var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                    IWorkbook workbook = new XSSFWorkbook(file);

                    var styles = new
                    {
                        FreeText = ExcelNPOIExtention.SetCellFreeStyle(workbook, true, HorizontalAlignment.Center, true, 16),
                        Text = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Left, true),
                        TextRight = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Right, true),
                        TextBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Left, true),
                        TextCenter = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Center, false),
                        TextCenterBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false),
                        Number = ExcelNPOIExtention.SetCellStyleNumber(workbook, false, HorizontalAlignment.Right, true),
                        NumberBold = ExcelNPOIExtention.SetCellStyleNumber(workbook, true, HorizontalAlignment.Right, true),
                    };
                    if (header.Date.Month > 4 && header.Date.Month < 11)
                    {
                        muaMien = "Hè Thu";
                        tuThang = "5 - 10";
                    }
                    else
                    {
                        muaMien = "Đông xuân";
                        tuThang = "11 - 4";

                    }
                    //if(accountGroup!= "G_NV_K") { 


                    #region Dữ liệu gốc
                    var sheetDlg = workbook.GetSheetAt(0);

                    #region thông báo
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(1) ?? sheetDlg.CreateRow(1), 6, $"THÔNG BÁO GIÁ BÁN LẺ TỪ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.FreeText);
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(2) ?? sheetDlg.CreateRow(2), 18, $"(Từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")})", ExcelNPOIExtention.SetCellFreeStyle(workbook, true, HorizontalAlignment.Center, false, 16));
                    int rowIndexDl1 = 5;
                    foreach (var i in data.Dlg.Dlg1)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl1) ?? sheetDlg.CreateRow(rowIndexDl1);
                        ExcelNPOIExtention.SetCellValueNumber(row, 1, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 2, i.Col2, number);
                        //ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueText(row, 4, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                    rowIndexDl1++;
                    }
                    int rowIndexDl2 = 5;
                    foreach (var i in data.Dlg.Dlg2)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl2) ?? sheetDlg.CreateRow(rowIndexDl2);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col2, number);
                        rowIndexDl2++;
                    }
                    int rowIndexDl3 = 5;
                    foreach (var i in data.Dlg.Dlg3)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl3) ?? sheetDlg.CreateRow(rowIndexDl3);
                        if (i.GoodCode == null)
                        {
                            ExcelNPOIExtention.SetCellValueText(row, 18, i.GoodName, styles.TextBold);
                        }
                        else
                        {
                            ExcelNPOIExtention.SetCellValueText(row, 18, i.GoodName, text);
                            ExcelNPOIExtention.SetCellValueNumber(row, 20, i.Col1, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                            //ExcelNPOIExtention.SetCellValueNumber(row, 24, i.Col3, number);
                            ExcelNPOIExtention.SetCellValueText(row, 24, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                    }
                    rowIndexDl3++;
                    }
                    #endregion

                    #region BIỂU TỔNG HỢP CÁC CHỈ TIÊU DẦU SÁNG (PT bán lẻ - V2)

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(34) ?? sheetDlg.CreateRow(34), 0, $"Tính từ: {header.Hour?.ToString("HH:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số {header.CongDienPtBanLe.ToString()} ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                    int rowIndexDl4 = 40;

                    #region xuất người ký
                    if (header.SignerCode == "TongGiamDoc")
                    {
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(46) ?? sheetDlg.CreateRow(46), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(47) ?? sheetDlg.CreateRow(93), 9, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(99), 9, $"{nguoiKy.Name}", styles.TextCenter);

                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(92) ?? sheetDlg.CreateRow(92), 11, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(93) ?? sheetDlg.CreateRow(93), 11, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 11, $"{nguoiKy.Name}", styles.TextCenter);

                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(128) ?? sheetDlg.CreateRow(128), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(129) ?? sheetDlg.CreateRow(129), 9, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 9, $"{nguoiKy.Name}", styles.TextCenter);

                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(164) ?? sheetDlg.CreateRow(164), 9, $"Vinh, Ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(165) ?? sheetDlg.CreateRow(165), 9, "GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 9, $"{nguoiKy.Name}", styles.TextCenter);
                    }
                    else
                    {
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(46) ?? sheetDlg.CreateRow(46), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(47) ?? sheetDlg.CreateRow(47), 9, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(48) ?? sheetDlg.CreateRow(48), 9, $"{nguoiKy.Position}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(53) ?? sheetDlg.CreateRow(53), 9, $"{nguoiKy.Name}", styles.TextCenter);

                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(92) ?? sheetDlg.CreateRow(92), 11, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(93) ?? sheetDlg.CreateRow(93), 11, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(94) ?? sheetDlg.CreateRow(94), 11, $"{nguoiKy.Position}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(99) ?? sheetDlg.CreateRow(99), 11, $"{nguoiKy.Name}", styles.TextCenter);

                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(128) ?? sheetDlg.CreateRow(128), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(129) ?? sheetDlg.CreateRow(129), 9, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(130) ?? sheetDlg.CreateRow(130), 9, $"{nguoiKy.Position}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(135) ?? sheetDlg.CreateRow(135), 9, $"{nguoiKy.Name}", styles.TextCenter);

                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(164) ?? sheetDlg.CreateRow(164), 9, $"Vinh, Ngày {header.Date.ToString("dd/ MM/ yyyy")}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(165) ?? sheetDlg.CreateRow(165), 9, "KT.GIÁM ĐỐC CÔNG TY", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(167) ?? sheetDlg.CreateRow(167), 9, $"{nguoiKy.Position}", styles.TextCenter);
                        ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(171) ?? sheetDlg.CreateRow(171), 9, $"{nguoiKy.Name}", styles.TextCenter);
                    }
                    #endregion

                    foreach (var i in data.Dlg.Dlg4)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl4) ?? sheetDlg.CreateRow(rowIndexDl4);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 2, i.Col1, styles.TextRight);
                        ExcelNPOIExtention.SetCellValueNumber(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col4, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col5, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col6, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col7, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col8, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col10, number);
                        rowIndexDl4++;
                    }
                    #endregion

                    #region THAY ĐỔI GIÁ GIAO PT BÁN LẺ
                    int rowIndexDl5 = 39;
                    foreach (var i in data.Dlg.Dlg3)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl5) ?? sheetDlg.CreateRow(rowIndexDl5);
                        if (i.GoodCode == null)
                        {
                            ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, styles.TextBold);
                        }
                        else
                        {
                            ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                            ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                            ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));
                            //ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);
                            ExcelNPOIExtention.SetCellValueText(row, 24, i.Col3 == 0 ? "(Không thay đổi)" : "(Thay đổi)", text);
                        }
                        rowIndexDl5++;
                    }
                    #endregion

                    #region BIỂU TỔNG HỢP CÁC CHỈ TIÊU DẦU SÁNG (ngoài bán lẻ)

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(74) ?? sheetDlg.CreateRow(74), 0, $"Tính từ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số  {header.CongDienSo.ToString()}   ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                    int rowIndexDl6 = 81;

                    foreach (var i in data.Dlg.Dlg6)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl6) ?? sheetDlg.CreateRow(rowIndexDl6);
                        if (i.GoodCode == null)
                        {
                            ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, styles.TextBold);
                        }
                        else
                        {
                            ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);

                            ExcelNPOIExtention.SetCellValueText(row, 2, i.Col1, styles.TextRight);
                            ExcelNPOIExtention.SetCellValueNumber(row, 3, i.Col2, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col4, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col5, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col6, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col7, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col8, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col9, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col10, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col11, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col12, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col13, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col14, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 16, i.Col15, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col16, number);
                        }
                        rowIndexDl6++;
                    }
                    #endregion

                    #region BIỂU TÍNH GIÁ XUẤT NỘI DỤNG

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(115) ?? sheetDlg.CreateRow(115), 0, $"Tính từ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số  {header.CongDienSo.ToString()}  ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                    int rowIndexDl7 = 121;
                    foreach (var i in data.Dlg.Dlg7)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl7) ?? sheetDlg.CreateRow(rowIndexDl7);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);

                        ExcelNPOIExtention.SetCellValueText(row, 3, i.Col1, styles.TextRight);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col4, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col5, number);

                        rowIndexDl7++;
                    }
                    #endregion

                    #region BIỂU TÍNH GIÁ BÁN CHO CÔNG TY CP VẬN TẢI VÀ DỊCH VỤ PETROLIMEX NGHỆ TĨNH

                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(153) ?? sheetDlg.CreateRow(153), 0, $"Tính từ: {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} theo CĐ số {header.CongDienSo.ToString()} ngày {header.Date.ToString("dd/MM/yyyy")}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {header.Date.ToString("dd/MM/yyyy")} và theo VCF {muaMien}", styles.TextCenter);
                    int rowIndexDl8 = 159;
                    if (header.SignerCode == "TongGiamDoc")
                        foreach (var i in data.Dlg.Dlg8)
                        {
                            var text = i.IsBold ? styles.TextBold : styles.Text;
                            var number = i.IsBold ? styles.NumberBold : styles.Number;
                            var row = sheetDlg.GetRow(rowIndexDl8) ?? sheetDlg.CreateRow(rowIndexDl8);
                            ExcelNPOIExtention.SetCellValueText(row, 1, i.GoodName, text);

                            ExcelNPOIExtention.SetCellValueText(row, 3, i.Col1, styles.TextRight);
                            ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col2, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col3, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col4, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col5, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col6, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col7, number);

                            rowIndexDl8++;
                        }
                    #endregion

                    #region SO SÁNH

                    #region thay đổi) giá giao phương thức bán lẻ
                    int rowIndexDl13 = 53;
                    var check = 159;
                    foreach (var i in data.Dlg.Dlg5)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl13) ?? sheetDlg.CreateRow(rowIndexDl13);
                        ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                        //ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));

                        ExcelNPOIExtention.SetCellValueText(sheetDlg.GetRow(check) ?? sheetDlg.CreateRow(check), 12, i.Col3 == 0 ? "(Không thay đổi)" : "(Thay đổi)", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Center, true, 13));
                        check++;
                        rowIndexDl13++;
                    }
                    #endregion

                    ExcelNPOIExtention.SetCellValueText(sheetDlg.GetRow(73) ?? sheetDlg.CreateRow(73), 20, $"{header.Date.ToString("dd.MM.yyyy")}", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Center, true, 13));
                    ExcelNPOIExtention.SetCellValueText(sheetDlg.GetRow(73) ?? sheetDlg.CreateRow(73), 21, $"{header.Date.ToString("hh:mm")}", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Center, true, 13));

                    #region thay đổi) lãi gộp
                    int rowIndexDl9_1 = 81;
                    int rowIndexDl9_2 = 86;
                    ExcelNPOIExtention.SetCellValue(sheetDlg.GetRow(76) ?? sheetDlg.CreateRow(76), 20, $"Lãi gộp từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")} và tính theo VCF {muaMien} từ tháng {tuThang} hàng năm", styles.TextCenter);

                    foreach (var i in data.Dlg.Dlg9)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        if (i.LocalCode == "V1")
                        {
                            var row = sheetDlg.GetRow(rowIndexDl9_1) ?? sheetDlg.CreateRow(rowIndexDl9_1);
                            ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                            ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);

                            rowIndexDl9_1++;
                        }
                        else
                        {
                            var row = sheetDlg.GetRow(rowIndexDl9_2) ?? sheetDlg.CreateRow(rowIndexDl9_2);

                            ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                            ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                            ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);

                            rowIndexDl9_2++;
                        }

                    }
                    #endregion

                    #region thay đổi chiết khấu
                    int rowIndexDl10 = 95;
                    foreach (var i in data.Dlg.Dlg10)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDlg.GetRow(rowIndexDl10) ?? sheetDlg.CreateRow(rowIndexDl10);
                        ExcelNPOIExtention.SetCellValueText(row, 20, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueText(row, 23, i.Col3 < 0 ? "(-) " + Math.Abs(i.Col3).ToString() : i.Col3 > 0 ? "(+) " + i.Col3 : "0", ExcelNPOIExtention.SetCellFreeStyle(workbook, false, HorizontalAlignment.Right, true, 13));
                        //ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col3, number);

                    rowIndexDl10++;
                    }
                    #endregion

                    #endregion

                    #endregion
                    //}

                    #region PT
                    var sheetPt = workbook.GetSheetAt(1);
                    ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(1) ?? sheetPt.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPt = 7;
                    foreach (var i in data.Pt)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPt.GetRow(rowIndexPt) ?? sheetPt.CreateRow(rowIndexPt);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col4, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 6, i.Col5, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 7, i.Col6, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 8, i.Col7, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col8, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col9, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col10, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col11, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col13, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 16, i.Col15, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col16, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 18, i.Col17, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 19, i.Col18, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 20, i.Col19, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 21, i.Col20, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 22, i.Col21, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 23, i.Col22, number);
                      
                        rowIndexPt++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 15, "DUYỆT", styles.TextCenterBold);
                    #endregion

                    #region ĐB
                    var sheetDb = workbook.GetSheetAt(2);
                    ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(1) ?? sheetDb.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexDb = 7;
                    foreach (var i in data.Db)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetDb.GetRow(rowIndexDb) ?? sheetDb.CreateRow(rowIndexDb);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValue(row, 2, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 6, i.Col4, number);
                        ExcelNPOIExtention.SetCellValue(row, 7, i.Col5, number);
                        ExcelNPOIExtention.SetCellValue(row, 8, i.Col6, number);
                        ExcelNPOIExtention.SetCellValue(row, 9, i.Col7, number);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                        ExcelNPOIExtention.SetCellValue(row, 15, i.Col13, number);
                        ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                        ExcelNPOIExtention.SetCellValue(row, 17, i.Col15, number);
                        ExcelNPOIExtention.SetCellValue(row, 18, i.Col16, number);
                        ExcelNPOIExtention.SetCellValue(row, 19, i.Col17, number);
                        ExcelNPOIExtention.SetCellValue(row, 20, i.Col18, number);
                        ExcelNPOIExtention.SetCellValue(row, 21, i.Col19, number);
                        ExcelNPOIExtention.SetCellValue(row, 22, i.Col20, number);
                        ExcelNPOIExtention.SetCellValue(row, 23, i.Col21, number);
                        ExcelNPOIExtention.SetCellValue(row, 24, i.Col22, number);
                        ExcelNPOIExtention.SetCellValue(row, 25, i.Col23, number);
                        ExcelNPOIExtention.SetCellValue(row, 26, i.Col24, number);
                        ExcelNPOIExtention.SetCellValue(row, 27, i.Col25, number);
                        ExcelNPOIExtention.SetCellValue(row, 28, i.Col26, number);
                  
                        rowIndexDb++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region FOB
                    var sheetFob = workbook.GetSheetAt(3);
                    ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(1) ?? sheetFob.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexFob = 7;
                    foreach (var i in data.Fob)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetFob.GetRow(rowIndexFob) ?? sheetFob.CreateRow(rowIndexFob);

                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);
                        ExcelNPOIExtention.SetCellValue(row, 6, i.Col5, number);
                        ExcelNPOIExtention.SetCellValue(row, 7, i.Col6, number);
                        ExcelNPOIExtention.SetCellValue(row, 8, i.Col7, number);
                        ExcelNPOIExtention.SetCellValue(row, 9, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col9, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col12, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col13, number);
                        ExcelNPOIExtention.SetCellValue(row, 15, i.Col14, number);
                        ExcelNPOIExtention.SetCellValue(row, 16, i.Col15, number);
                        ExcelNPOIExtention.SetCellValue(row, 17, i.Col16, number);
                        ExcelNPOIExtention.SetCellValue(row, 18, i.Col17, number);
                        ExcelNPOIExtention.SetCellValue(row, 19, i.Col18, number);
                        ExcelNPOIExtention.SetCellValue(row, 20, i.Col19, number);
                        ExcelNPOIExtention.SetCellValue(row, 21, i.Col20, number);
                        ExcelNPOIExtention.SetCellValue(row, 22, i.Col21, number);
                        ExcelNPOIExtention.SetCellValue(row, 23, i.Col22, number);
                        ExcelNPOIExtention.SetCellValue(row, 24, i.Col23, number);
                        ExcelNPOIExtention.SetCellValue(row, 25, i.Col24, number);
                   

                        rowIndexFob++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region PT 09
                    var sheetPt09 = workbook.GetSheetAt(4);
                    ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(1) ?? sheetPt09.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPt09 = 7;
                    foreach (var i in data.Pt09)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPt09.GetRow(rowIndexPt09) ?? sheetPt09.CreateRow(rowIndexPt09);

                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);
                        ExcelNPOIExtention.SetCellValue(row, 6, i.Col5, number);
                        ExcelNPOIExtention.SetCellValue(row, 7, i.Col6, number);
                        ExcelNPOIExtention.SetCellValue(row, 8, i.Col7, number);
                        ExcelNPOIExtention.SetCellValue(row, 9, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col12, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col13, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col14, number);
                        ExcelNPOIExtention.SetCellValue(row, 15, i.Col15, number);
                        ExcelNPOIExtention.SetCellValue(row, 16, i.Col16, number);
                        ExcelNPOIExtention.SetCellValue(row, 17, i.Col17, number);
                        ExcelNPOIExtention.SetCellValue(row, 18, i.Col18, number);
                        ExcelNPOIExtention.SetCellValue(row, 19, i.Col19, number);
                        ExcelNPOIExtention.SetCellValue(row, 20, i.Col20, number);
                        ExcelNPOIExtention.SetCellValue(row, 21, i.Col21, number);
                        ExcelNPOIExtention.SetCellValue(row, 22, i.Col22, number);
                        ExcelNPOIExtention.SetCellValue(row, 23, i.Col23, number);
                        ExcelNPOIExtention.SetCellValue(row, 24, i.Col24, number);
                    

                        rowIndexPt09++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region BB DO
                    var sheetBbDo = workbook.GetSheetAt(5);
                    ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(3) ?? sheetBbDo.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexBbDo = 9;
                    foreach (var i in data.Bbdo)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetBbDo.GetRow(rowIndexBbDo) ?? sheetBbDo.CreateRow(rowIndexBbDo);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.DeliveryPoint, text);
                        ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 4, i.PThuc, text);
                        ExcelNPOIExtention.SetCellValueText(row, 5, i.CustomerCode, text);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.GoodCode, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Dvt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 8, i.TToan, text);
                        ExcelNPOIExtention.SetCellValue(row, 9, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col4, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col5, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col6, number);
                        ExcelNPOIExtention.SetCellValue(row, 15, i.Col7, number);
                        ExcelNPOIExtention.SetCellValue(row, 16, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 17, i.Col9, number);
                        ExcelNPOIExtention.SetCellValue(row, 18, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 19, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 20, i.Col12, number);
                        ExcelNPOIExtention.SetCellValue(row, 21, i.Col13, number);
                        ExcelNPOIExtention.SetCellValue(row, 22, i.Col14, number);
                        rowIndexBbDo++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region BB FO

                    #endregion

                    #region PL1
                    var sheetPl1 = workbook.GetSheetAt(7);
                    ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(2) ?? sheetPl1.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPl1 = 8;
                    foreach (var i in data.Pl1)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPl1.GetRow(rowIndexPl1) ?? sheetPl1.CreateRow(rowIndexPl1);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.MarketName, text);

                        ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                        rowIndexPl1++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region PL2

                    var sheetPl2 = workbook.GetSheetAt(8);
                    ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(2) ?? sheetPl2.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPl2 = 7;
                    foreach (var i in data.Pl2)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPl2.GetRow(rowIndexPl2) ?? sheetPl2.CreateRow(rowIndexPl2);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                        ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                        rowIndexPl2++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region PL3

                    var sheetPl3 = workbook.GetSheetAt(9);
                    ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(2) ?? sheetPl3.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPl3 = 7;
                    foreach (var i in data.Pl3)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPl3.GetRow(rowIndexPl3) ?? sheetPl3.CreateRow(rowIndexPl3);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                        ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                        rowIndexPl3++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region PL4

                    var sheetPl4 = workbook.GetSheetAt(10);
                    ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(3) ?? sheetPl4.CreateRow(3), 2, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPl4 = 8;
                    foreach (var i in data.Pl4)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPl4.GetRow(rowIndexPl4) ?? sheetPl4.CreateRow(rowIndexPl4);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                        ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                        ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                        rowIndexPl4++;
                    }
                    ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                    ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 15, "DUYỆT", styles.TextCenterBold);

                    #endregion

                    #region VK11-PT

                    var sheetVk11Pt = workbook.GetSheetAt(11);
                    //ExcelNPOIExtention.SetCellValue(sheetVk11Pt.GetRow(3) ?? sheetVk11Pt.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexVk11Pt = 2;
                    foreach (var i in data.Vk11Pt)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetVk11Pt.GetRow(rowIndexVk11Pt) ?? sheetVk11Pt.CreateRow(rowIndexVk11Pt);
                        ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                        ExcelNPOIExtention.SetCellValueText(row, 3, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                        ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                        ExcelNPOIExtention.SetCellValueText(row, 9, i.Col6, text);
                        ExcelNPOIExtention.SetCellValueText(row, 10, i.Col7, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                        ExcelNPOIExtention.SetCellValueText(row, 12, i.Col9, text);
                        ExcelNPOIExtention.SetCellValueText(row, 13, i.Col10, text);
                        ExcelNPOIExtention.SetCellValueText(row, 14, i.Col11, text);
                        ExcelNPOIExtention.SetCellValueText(row, 15, i.Col12, text);
                        ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                        ExcelNPOIExtention.SetCellValueText(row, 17, i.Col14, text);
                        ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                        rowIndexVk11Pt++;
                    }
                    #endregion

                    #region VK11-ĐB

                    var sheetVk11Db = workbook.GetSheetAt(12);
                    //ExcelNPOIExtention.SetCellValue(sheetVk11Db.GetRow(3) ?? sheetVk11Db.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexVk11Db = 2;
                    foreach (var i in data.Vk11Db)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetVk11Db.GetRow(rowIndexVk11Db) ?? sheetVk11Db.CreateRow(rowIndexVk11Db);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValue(row, 2, i.Address, text);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 4, i.Col1, text);
                        ExcelNPOIExtention.SetCellValueText(row, 5, i.Col2, text);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                        ExcelNPOIExtention.SetCellValue(row, 8, i.Col5, number);
                        ExcelNPOIExtention.SetCellValue(row, 9, i.Col6, number);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col7, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col9, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 15, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                        ExcelNPOIExtention.SetCellValue(row, 17, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                        rowIndexVk11Db++;
                    }
                    #endregion

                    #region VK11-FOB

                    var sheetVk11Fob = workbook.GetSheetAt(13);
                    //ExcelNPOIExtention.SetCellValue(sheetVk11Fob.GetRow(3) ?? sheetVk11Fob.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexVk11Fob = 3;
                    foreach (var i in data.Vk11Fob)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetVk11Fob.GetRow(rowIndexVk11Fob) ?? sheetVk11Fob.CreateRow(rowIndexVk11Fob);
                        ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 3, "", number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 4, "", number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, "", number);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                        ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                        ExcelNPOIExtention.SetCellValueText(row, 9, i.Col6, text);
                        ExcelNPOIExtention.SetCellValueText(row, 10, i.Col7, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col10, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col11, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                        rowIndexVk11Fob++;
                    }
                    #endregion

                    #region VK11-TNPP

                    var sheetVk11Tnpp = workbook.GetSheetAt(14);
                    //ExcelNPOIExtention.SetCellValue(sheetVk11Tnpp.GetRow(3) ?? sheetVk11Tnpp.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexVk11Tnpp = 3;
                    foreach (var i in data.Vk11Tnpp)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetVk11Tnpp.GetRow(rowIndexVk11Tnpp) ?? sheetVk11Tnpp.CreateRow(rowIndexVk11Tnpp);
                        ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                        ExcelNPOIExtention.SetCellValue(row, 3, i.Col1, number);
                        ExcelNPOIExtention.SetCellValue(row, 4, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueText(row, 5, i.Col3, text);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.Col4, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Col5, text);
                        ExcelNPOIExtention.SetCellValueText(row, 8, i.Col6, text);
                        ExcelNPOIExtention.SetCellValueText(row, 9, i.Col7, text);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueText(row, 15, i.Col13, text);
                        ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueText(row, 17, i.Col15, text);
                        rowIndexVk11Tnpp++;
                    }
                    #endregion

                    #region VK11-BB

                    var sheetVk11Bb = workbook.GetSheetAt(15);
                    //ExcelNPOIExtention.SetCellValue(sheetVk11Bb.GetRow(3) ?? sheetVk11Bb.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexVk11Bb = 2;
                    foreach (var i in data.Vk11Bb)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetVk11Bb.GetRow(rowIndexVk11Bb) ?? sheetVk11Bb.CreateRow(rowIndexVk11Bb);
                        ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                        ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodsName, text);
                        ExcelNPOIExtention.SetCellValue(row, 4, "", number);
                        ExcelNPOIExtention.SetCellValueText(row, 5, i.Col3, text);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.Col4, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Col5, text);
                        ExcelNPOIExtention.SetCellValue(row, 8, i.Col6, number);
                        ExcelNPOIExtention.SetCellValue(row, 9, i.Col7, number);
                        ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                        ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                        ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                        ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                        ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueText(row, 15, i.Col13, text);
                        ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueText(row, 17, i.Col15, text);
                        rowIndexVk11Bb++;
                    }
                    #endregion

                    #region PTS
                    var sheetPts = workbook.GetSheetAt(16);
                    //ExcelNPOIExtention.SetCellValue(sheetVk11Pt.GetRow(3) ?? sheetVk11Pt.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexPts = 4;
                    foreach (var i in data.Pts)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetPts.GetRow(rowIndexPts) ?? sheetPts.CreateRow(rowIndexPts);
                        ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.PThuc, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.CustomerCode, text);
                        ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodCode, text);
                        ExcelNPOIExtention.SetCellValueText(row, 4, "L15", text);
                        ExcelNPOIExtention.SetCellValueText(row, 5, i.TToan, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 6, "", number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 7, "", number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 8, "", number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col3, number);
                        ExcelNPOIExtention.SetCellValueText(row, 10, "", text);
                        ExcelNPOIExtention.SetCellValueText(row, 11, "", text);
                        ExcelNPOIExtention.SetCellValueText(row, 12, "L15", text);
                        ExcelNPOIExtention.SetCellValueText(row, 13, "", text);
                        ExcelNPOIExtention.SetCellValueText(row, 14, "", text);
                        ExcelNPOIExtention.SetCellValueText(row, 15, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 16, "31.12.9999", text);
                        ExcelNPOIExtention.SetCellValueText(row, 17, "", text);
                        ExcelNPOIExtention.SetCellValueText(row, 18, "", text);
                        ExcelNPOIExtention.SetCellValueText(row, 19, "PTS", text);

                        rowIndexPts++;
                    }
                    #endregion

                    #region Tổng hợp

                    var sheetTh = workbook.GetSheetAt(17);
                    //ExcelNPOIExtention.SetCellValue(sheetTh.GetRow(3) ?? sheetTh.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                    int rowIndexTh = 2;
                    foreach (var i in data.Summary)
                    {
                        var text = i.IsBold ? styles.TextBold : styles.Text;
                        var number = i.IsBold ? styles.NumberBold : styles.Number;
                        var row = sheetTh.GetRow(rowIndexTh) ?? sheetTh.CreateRow(rowIndexTh);

                        ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                        ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                        ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                        ExcelNPOIExtention.SetCellValueText(row, 3, i.MarketName, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col1, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                        ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                        ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                        ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col6, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col7, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col10, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col11, number);
                        ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col12, number);
                        ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                        ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col14, number);
                        ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                        rowIndexTh++;
                    }
                    #endregion




                    var folderPath = Path.Combine($"Uploads/Excel/{DateTime.Now.ToString("yyyy/MM/dd")}");
                    if (!Directory.Exists(folderPath))
                        Directory.CreateDirectory(folderPath);
                    var fileName = $"CSTMGG_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                    var outputPath = Path.Combine(Directory.GetCurrentDirectory(), folderPath, fileName);

                    using var outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                    workbook.Write(outFile);

                    _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = headerId,
                        Name = fileName,
                        Type = "xlsx",
                        Path = $"{folderPath}/{fileName}",
                    });
                    await _dbContext.SaveChangesAsync();

                    return $"{folderPath}/{fileName}";
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Lỗi khi xuất file Excel: {ex.Message}\n{ex.StackTrace}");
                    return string.Empty;
                }
            }
        
        #endregion

        #region Xuất Excel NVK (Tất cả sheet )
        public async Task<string> ExportExcelNVK(string headerId)
        {
            try
            {
                var header = _dbContext.TblBuCalculateDiscount.Find(headerId);
                var nguoiKy = _dbContext.TblMdSigner.Where(x => x.Code == header.SignerCode).FirstOrDefault();
                var data = await this.CalculateDiscountOutput(headerId);
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template", "CoSoTinhMucGiamGiaNVK.xlsx");
                var muaMien = "";
                var tuThang = "";

                if (!File.Exists(templatePath))
                {
                    throw new FileNotFoundException("Không tìm thấy template Excel.", templatePath);
                }

                using var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read, FileShare.Read);
                IWorkbook workbook = new XSSFWorkbook(file);

                var styles = new
                {
                    FreeText = ExcelNPOIExtention.SetCellFreeStyle(workbook, true, HorizontalAlignment.Center, true, 16),
                    Text = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Left, true),
                    TextRight = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Right, true),
                    TextBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Left, true),
                    TextCenter = ExcelNPOIExtention.SetCellStyleText(workbook, false, HorizontalAlignment.Center, false),
                    TextCenterBold = ExcelNPOIExtention.SetCellStyleText(workbook, true, HorizontalAlignment.Center, false),
                    Number = ExcelNPOIExtention.SetCellStyleNumber(workbook, false, HorizontalAlignment.Right, true),
                    NumberBold = ExcelNPOIExtention.SetCellStyleNumber(workbook, true, HorizontalAlignment.Right, true),
                };
                if (header.Date.Month > 4 && header.Date.Month < 11)
                {
                    muaMien = "Hè Thu";
                    tuThang = "5 - 10";
                }
                else
                {
                    muaMien = "Đông xuân";
                    tuThang = "11 - 4";

                }
                //if(accountGroup!= "G_NV_K") { 


              
                //}

                #region PT
                var sheetPt = workbook.GetSheetAt(0);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(1) ?? sheetPt.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPt = 7;
                foreach (var i in data.Pt)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPt.GetRow(rowIndexPt) ?? sheetPt.CreateRow(rowIndexPt);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.MarketName, text);
                   
                    ExcelNPOIExtention.SetCellValueNumber(row, 24, i.Col23, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 25, i.Col24, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 26, i.Col25, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 27, i.Col26, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 28, i.Col27, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 29, i.Col28, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 30, i.Col29, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 31, i.Col30, number);
                    rowIndexPt++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt.GetRow(rowIndexPt + 1) ?? sheetPt.CreateRow(rowIndexPt + 1), 15, "DUYỆT", styles.TextCenterBold);
                #endregion

                #region ĐB
                var sheetDb = workbook.GetSheetAt(1);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(1) ?? sheetDb.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexDb = 7;
                foreach (var i in data.Db)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetDb.GetRow(rowIndexDb) ?? sheetDb.CreateRow(rowIndexDb);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValue(row, 2, i.MarketName, text);
                   
                    ExcelNPOIExtention.SetCellValue(row, 29, i.Col27, number);
                    ExcelNPOIExtention.SetCellValue(row, 30, i.Col28, number);
                    ExcelNPOIExtention.SetCellValue(row, 31, i.Col29, number);
                    ExcelNPOIExtention.SetCellValue(row, 32, i.Col30, number);
                    ExcelNPOIExtention.SetCellValue(row, 33, i.Col31, number);
                    ExcelNPOIExtention.SetCellValue(row, 34, i.Col32, number);
                    ExcelNPOIExtention.SetCellValue(row, 35, i.Col33, number);
                    ExcelNPOIExtention.SetCellValue(row, 36, i.Col34, number);
                    rowIndexDb++;
                }
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetDb.GetRow(rowIndexDb + 1) ?? sheetDb.CreateRow(rowIndexDb + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region FOB
                var sheetFob = workbook.GetSheetAt(2);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(1) ?? sheetFob.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexFob = 7;
                foreach (var i in data.Fob)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetFob.GetRow(rowIndexFob) ?? sheetFob.CreateRow(rowIndexFob);

                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
               
                    ExcelNPOIExtention.SetCellValue(row, 26, i.Col25, number);
                    ExcelNPOIExtention.SetCellValue(row, 27, i.Col26, number);
                    ExcelNPOIExtention.SetCellValue(row, 28, i.Col27, number);
                    ExcelNPOIExtention.SetCellValue(row, 29, i.Col28, number);
                    ExcelNPOIExtention.SetCellValue(row, 30, i.Col29, number);
                    ExcelNPOIExtention.SetCellValue(row, 31, i.Col30, number);
                    ExcelNPOIExtention.SetCellValue(row, 32, i.Col31, number);
                    ExcelNPOIExtention.SetCellValue(row, 33, i.Col32, number);

                    rowIndexFob++;
                }
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetFob.GetRow(rowIndexFob + 1) ?? sheetFob.CreateRow(rowIndexFob + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region PT 09
                var sheetPt09 = workbook.GetSheetAt(3);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(1) ?? sheetPt09.CreateRow(1), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPt09 = 7;
                foreach (var i in data.Pt09)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPt09.GetRow(rowIndexPt09) ?? sheetPt09.CreateRow(rowIndexPt09);

                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                  
                    ExcelNPOIExtention.SetCellValue(row, 25, i.Col25, number);
                    ExcelNPOIExtention.SetCellValue(row, 26, i.Col26, number);
                    ExcelNPOIExtention.SetCellValue(row, 27, i.Col27, number);
                    ExcelNPOIExtention.SetCellValue(row, 28, i.Col28, number);
                    ExcelNPOIExtention.SetCellValue(row, 29, i.Col29, number);
                    ExcelNPOIExtention.SetCellValue(row, 30, i.Col30, number);
                    ExcelNPOIExtention.SetCellValue(row, 31, i.Col31, number);
                    ExcelNPOIExtention.SetCellValue(row, 32, i.Col32, number);
                    //ExcelNPOIExtention.SetCellValue(row, 33, i.Col33, number);

                    rowIndexPt09++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPt09.GetRow(rowIndexPt09 + 1) ?? sheetPt09.CreateRow(rowIndexPt09 + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region BB DO
                var sheetBbDo = workbook.GetSheetAt(4);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(3) ?? sheetBbDo.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexBbDo = 9;
                foreach (var i in data.Bbdo)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetBbDo.GetRow(rowIndexBbDo) ?? sheetBbDo.CreateRow(rowIndexBbDo);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.DeliveryPoint, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 4, i.PThuc, text);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.CustomerCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.GoodCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Dvt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.TToan, text);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col4, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 18, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 19, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 20, i.Col12, number);
                    ExcelNPOIExtention.SetCellValue(row, 21, i.Col13, number);
                    ExcelNPOIExtention.SetCellValue(row, 22, i.Col14, number);
                    rowIndexBbDo++;
                }
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetBbDo.GetRow(rowIndexBbDo + 1) ?? sheetBbDo.CreateRow(rowIndexBbDo + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region BB FO

                #endregion

                #region PL1
                var sheetPl1 = workbook.GetSheetAt(6);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(2) ?? sheetPl1.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl1 = 8;
                foreach (var i in data.Pl1)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl1.GetRow(rowIndexPl1) ?? sheetPl1.CreateRow(rowIndexPl1);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.MarketName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl1++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl1.GetRow(rowIndexPl1 + 1) ?? sheetPl1.CreateRow(rowIndexPl1 + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region PL2

                var sheetPl2 = workbook.GetSheetAt(7);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(2) ?? sheetPl2.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl2 = 7;
                foreach (var i in data.Pl2)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl2.GetRow(rowIndexPl2) ?? sheetPl2.CreateRow(rowIndexPl2);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl2++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl2.GetRow(rowIndexPl2 + 1) ?? sheetPl2.CreateRow(rowIndexPl2 + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region PL3

                var sheetPl3 = workbook.GetSheetAt(8);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(2) ?? sheetPl3.CreateRow(2), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl3 = 7;
                foreach (var i in data.Pl3)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl3.GetRow(rowIndexPl3) ?? sheetPl3.CreateRow(rowIndexPl3);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl3++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl3.GetRow(rowIndexPl3 + 1) ?? sheetPl3.CreateRow(rowIndexPl3 + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region PL4

                var sheetPl4 = workbook.GetSheetAt(9);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(3) ?? sheetPl4.CreateRow(3), 2, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPl4 = 8;
                foreach (var i in data.Pl4)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPl4.GetRow(rowIndexPl4) ?? sheetPl4.CreateRow(rowIndexPl4);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);

                    ExcelNPOIExtention.SetCellValue(row, 2, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col2, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col3, number);
                    ExcelNPOIExtention.SetCellValue(row, 5, i.Col4, number);

                    rowIndexPl4++;
                }
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 1, "LẬP BIỂU", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 5, "P. KINH DOANH XD", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 10, "PHÒNG TCKT", styles.TextCenterBold);
                ExcelNPOIExtention.SetCellValue(sheetPl4.GetRow(rowIndexPl4 + 1) ?? sheetPl4.CreateRow(rowIndexPl4 + 1), 15, "DUYỆT", styles.TextCenterBold);

                #endregion

                #region VK11-PT

                var sheetVk11Pt = workbook.GetSheetAt(10);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Pt.GetRow(3) ?? sheetVk11Pt.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Pt = 2;
                foreach (var i in data.Vk11Pt)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Pt.GetRow(rowIndexVk11Pt) ?? sheetVk11Pt.CreateRow(rowIndexVk11Pt);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueText(row, 9, i.Col6, text);
                    ExcelNPOIExtention.SetCellValueText(row, 10, i.Col7, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueText(row, 12, i.Col9, text);
                    ExcelNPOIExtention.SetCellValueText(row, 13, i.Col10, text);
                    ExcelNPOIExtention.SetCellValueText(row, 14, i.Col11, text);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.Col12, text);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValueText(row, 17, i.Col14, text);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexVk11Pt++;
                }
                #endregion

                #region VK11-ĐB

                var sheetVk11Db = workbook.GetSheetAt(11);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Db.GetRow(3) ?? sheetVk11Db.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Db = 2;
                foreach (var i in data.Vk11Db)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Db.GetRow(rowIndexVk11Db) ?? sheetVk11Db.CreateRow(rowIndexVk11Db);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValue(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValue(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 4, i.Col1, text);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.Col2, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col5, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 15, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValue(row, 17, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexVk11Db++;
                }
                #endregion

                #region VK11-FOB

                var sheetVk11Fob = workbook.GetSheetAt(12);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Fob.GetRow(3) ?? sheetVk11Fob.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Fob = 3;
                foreach (var i in data.Vk11Fob)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Fob.GetRow(rowIndexVk11Fob) ?? sheetVk11Fob.CreateRow(rowIndexVk11Fob);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 3, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, "", number);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueText(row, 9, i.Col6, text);
                    ExcelNPOIExtention.SetCellValueText(row, 10, i.Col7, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col10, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col11, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexVk11Fob++;
                }
                #endregion

                #region VK11-TNPP

                var sheetVk11Tnpp = workbook.GetSheetAt(13);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Tnpp.GetRow(3) ?? sheetVk11Tnpp.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Tnpp = 3;
                foreach (var i in data.Vk11Tnpp)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Tnpp.GetRow(rowIndexVk11Tnpp) ?? sheetVk11Tnpp.CreateRow(rowIndexVk11Tnpp);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValue(row, 3, i.Col1, number);
                    ExcelNPOIExtention.SetCellValue(row, 4, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col6, text);
                    ExcelNPOIExtention.SetCellValueText(row, 9, i.Col7, text);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.Col13, text);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 17, i.Col15, text);
                    rowIndexVk11Tnpp++;
                }
                #endregion

                #region VK11-BB

                var sheetVk11Bb = workbook.GetSheetAt(14);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Bb.GetRow(3) ?? sheetVk11Bb.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexVk11Bb = 2;
                foreach (var i in data.Vk11Bb)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetVk11Bb.GetRow(rowIndexVk11Bb) ?? sheetVk11Bb.CreateRow(rowIndexVk11Bb);
                    ExcelNPOIExtention.SetCellValue(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodsName, text);
                    ExcelNPOIExtention.SetCellValue(row, 4, "", number);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col5, text);
                    ExcelNPOIExtention.SetCellValue(row, 8, i.Col6, number);
                    ExcelNPOIExtention.SetCellValue(row, 9, i.Col7, number);
                    ExcelNPOIExtention.SetCellValue(row, 10, i.Col8, number);
                    ExcelNPOIExtention.SetCellValue(row, 11, i.Col9, number);
                    ExcelNPOIExtention.SetCellValue(row, 12, i.Col10, number);
                    ExcelNPOIExtention.SetCellValue(row, 13, i.Col11, number);
                    ExcelNPOIExtention.SetCellValue(row, 14, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.Col13, text);
                    ExcelNPOIExtention.SetCellValue(row, 16, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 17, i.Col15, text);
                    rowIndexVk11Bb++;
                }
                #endregion

                #region PTS
                var sheetPts = workbook.GetSheetAt(15);
                //ExcelNPOIExtention.SetCellValue(sheetVk11Pt.GetRow(3) ?? sheetVk11Pt.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexPts = 4;
                foreach (var i in data.Pts)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetPts.GetRow(rowIndexPts) ?? sheetPts.CreateRow(rowIndexPts);
                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.PThuc, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.CustomerCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.GoodCode, text);
                    ExcelNPOIExtention.SetCellValueText(row, 4, "L15", text);
                    ExcelNPOIExtention.SetCellValueText(row, 5, i.TToan, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 6, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 7, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 8, "", number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col3, number);
                    ExcelNPOIExtention.SetCellValueText(row, 10, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 11, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 12, "L15", text);
                    ExcelNPOIExtention.SetCellValueText(row, 13, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 14, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 15, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 16, "31.12.9999", text);
                    ExcelNPOIExtention.SetCellValueText(row, 17, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 18, "", text);
                    ExcelNPOIExtention.SetCellValueText(row, 19, "PTS", text);

                    rowIndexPts++;
                }
                #endregion

                #region Tổng hợp

                var sheetTh = workbook.GetSheetAt(16);
                //ExcelNPOIExtention.SetCellValue(sheetTh.GetRow(3) ?? sheetTh.CreateRow(3), 0, $"Thực hiện: từ {header.Date.ToString("hh:mm")} ngày {header.Date.ToString("dd/MM/yyyy")}", styles.TextCenter);
                int rowIndexTh = 2;
                foreach (var i in data.Summary)
                {
                    var text = i.IsBold ? styles.TextBold : styles.Text;
                    var number = i.IsBold ? styles.NumberBold : styles.Number;
                    var row = sheetTh.GetRow(rowIndexTh) ?? sheetTh.CreateRow(rowIndexTh);

                    ExcelNPOIExtention.SetCellValueText(row, 0, i.Stt, text);
                    ExcelNPOIExtention.SetCellValueText(row, 1, i.CustomerName, text);
                    ExcelNPOIExtention.SetCellValueText(row, 2, i.Address, text);
                    ExcelNPOIExtention.SetCellValueText(row, 3, i.MarketName, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 4, i.Col1, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 5, i.Col2, number);
                    ExcelNPOIExtention.SetCellValueText(row, 6, i.Col3, text);
                    ExcelNPOIExtention.SetCellValueText(row, 7, i.Col4, text);
                    ExcelNPOIExtention.SetCellValueText(row, 8, i.Col5, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 9, i.Col6, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 10, i.Col7, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 11, i.Col8, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 12, i.Col9, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 13, i.Col10, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 14, i.Col11, number);
                    ExcelNPOIExtention.SetCellValueNumber(row, 15, i.Col12, number);
                    ExcelNPOIExtention.SetCellValueText(row, 16, i.Col13, text);
                    ExcelNPOIExtention.SetCellValueNumber(row, 17, i.Col14, number);
                    ExcelNPOIExtention.SetCellValueText(row, 18, i.Col15, text);
                    rowIndexTh++;
                }
                #endregion




                var folderPath = Path.Combine($"Uploads/Excel/{DateTime.Now.ToString("yyyy/MM/dd")}");
                if (!Directory.Exists(folderPath))
                    Directory.CreateDirectory(folderPath);
                var fileName = $"CSTMGG_{DateTime.Now:ddMMyyyy_HHmmss}.xlsx";
                var outputPath = Path.Combine(Directory.GetCurrentDirectory(), folderPath, fileName);

                using var outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFile);

                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = fileName,
                    Type = "xlsx",
                    Path = $"{folderPath}/{fileName}",
                });
                await _dbContext.SaveChangesAsync();

                return $"{folderPath}/{fileName}";
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi khi xuất file Excel: {ex.Message}\n{ex.StackTrace}");
                return string.Empty;
            }
        }
        #endregion

        #region Xuất trình ký Excel
        public async Task<string> ExportExcelTrinhKy(string headerId)
        {
            try
            {
                var data = await CalculateDiscountOutput(headerId);
                //var header = await _dbContext.TblBuCalculateDiscount.FindAsync(headerId);
                var goods = await _dbContext.TblMdGoods.ToListAsync();
                var NguoiKyTen = await _dbContext.TblMdSigner.FirstOrDefaultAsync(x => x.Code == data.Header.SignerCode);
                var date = $"{data.Header.Date.Hour:D2}h{data.Header.Date.Minute:D2} ngày {data.Header.Date.ToString("dd/MM/yyyy")}"; 
                var A5 = $"  (Kèm theo Công văn số:                        /PLXNA ngày {data.Header.Date.Day:D2}/{data.Header.Date.Month:D2}/{data.Header.Date.Year} của Công ty Xăng dầu Nghệ An)";
                var A29 = $" + Căn cứ Quyết định số {data.Header.QuyetDinhSo} ngày {data.Header.Date.Day:D2}/{data.Header.Date.Month:D2}/{data.Header.Date.Year} của Tổng giám đốc Tập đoàn Xăng dầu Việt Nam về việc qui định giá bán xăng dầu; ";
                var B25 = $"Mức giá bán đăng ký này có hiệu lực thi hành kể từ {data.Header.Date.Hour:D2} giờ 00 ngày {data.Header.Date.Day:D2} tháng {data.Header.Date.Month:D2} năm {data.Header.Date.Year}";
                var E8 = $"Giá kê khai kì liền kề trước (theo VB số ......./PLXNA-KDXD ngày {data.HeaderOld.Date.ToString("dd/MM/yyyy")}";
                // 1. Đường dẫn file gốc
                var filePathTemplate = Path.Combine(Directory.GetCurrentDirectory(), "Template", "TempTrinhKy", "KeKhaiGiaChiTiet.xlsx");

                // 2. Tạo thư mục lưu file
                var folderName = Path.Combine($"Uploads/Excel/{DateTime.Now.ToString("yyyy/MM/dd")}");
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }

                // 3. Tạo tên file mới
                var fileName = $"{DateTime.Now:ddMMyyyy_HHmmss}_KeKhaiGiaChiTiet.xlsx";
                var fullPath = Path.Combine(folderName, fileName);

                // 4. Copy file từ Template sang Upload
                File.Copy(filePathTemplate, fullPath, true);

                // 5. Mở file để sửa
                IWorkbook workbook;

                using (var fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
                {
                    workbook = new XSSFWorkbook(fs);
                    ISheet sheet = workbook.GetSheetAt(0);

                    IRow rowA5 = sheet.GetRow(4);
                    ICell cellA5 = rowA5?.GetCell(0);

                    if (cellA5 != null)
                    {
                        cellA5.SetCellValue(A5);
                    }

                    IRow rowA29 = sheet.GetRow(28);
                    ICell cellA29 = rowA29?.GetCell(0);

                    if (cellA29 != null)
                    {
                        cellA29.SetCellValue(A29);
                    }

                    IRow rowE8 = sheet.GetRow(7);
                    ICell cellE8 = rowE8?.GetCell(4);
                    cellE8.SetCellValue(E8);

                    int rowIndex = 10; // Bắt đầu từ row 11 (index = 10)
                    foreach (var item in data.Dlg.Dlg3)
                    {
                        IRow row = sheet.GetRow(rowIndex); // Chỉ lấy row, không cần CreateRow
                        if (row != null && item.LocalCode == "V1" && item.IsBold == false)
                        {
                            // E11 -> col1
                            ICell cellE = row.GetCell(4);
                            if (cellE != null)
                            {
                                cellE.SetCellValue((double)item.Col1);
                            }

                            // F11 -> col2
                            ICell cellF = row.GetCell(5);
                            if (cellF != null)
                            {
                                cellF.SetCellValue((double)item.Col2);
                            }

                            //G11 
                            ICell cellG = row.GetCell(6);
                            if (cellG != null)
                            {
                                cellG.SetCellValue(date);
                            }

                            // H11 -> tangGiam1_2
                            ICell cellH = row.GetCell(7);
                            if (cellH != null)
                            {
                                cellH.SetCellValue((double)item.Col3);
                            }

                            ICell cellI = row.GetCell(8);
                            if (cellI != null)
                            {
                                if (item.Col1 != 0)
                                {
                                    double rateOfIncreaseAndDecrease = (double)((item.Col2 - item.Col1) / item.Col1);
                                    cellI.SetCellValue(rateOfIncreaseAndDecrease);
                                }
                                else
                                {
                                    cellI.SetCellValue(0);
                                }
                            }
                            rowIndex++;
                        }
                    }

                    int rowIndex2 = 15;
                    foreach (var item in data.Dlg.Dlg3)
                    {
                        IRow row = sheet.GetRow(rowIndex2); // Chỉ lấy row, không cần CreateRow
                        if (row != null && item.LocalCode == "V2" && item.IsBold == false)
                        {
                            //// B11 -> colA
                            //ICell cellB = row.GetCell(1);
                            //if (cellB != null)
                            //{
                            //    cellB.SetCellValue(item.GoodName);
                            //}

                            // E11 -> col1
                            ICell cellE = row.GetCell(4);
                            if (cellE != null)
                            {
                                cellE.SetCellValue((double)item.Col1);
                            }

                            // F11 -> col2
                            ICell cellF = row.GetCell(5);
                            if (cellF != null)
                            {
                                cellF.SetCellValue((double)item.Col2);
                            }

                            // G11
                            ICell cellG = row.GetCell(6);
                            if (cellG != null)
                            {
                                cellG.SetCellValue(date);
                            }

                            // G12 -> tangGiam1_2
                            ICell cellH = row.GetCell(7);
                            if (cellH != null)
                            {
                                cellH.SetCellValue((double)item.Col3);
                            }

                            ICell cellI = row.GetCell(8);
                            if (cellI != null)
                            {
                                if (item.Col1 != 0)
                                {
                                    double rateOfIncreaseAndDecrease = (double)((item.Col2 - item.Col1) / item.Col1);
                                    cellI.SetCellValue(rateOfIncreaseAndDecrease);
                                }
                                else
                                {
                                    cellH.SetCellValue(0);
                                }
                            }
                            rowIndex2++;
                        }

                    }
                    
                    int rowIndex3 = 20;
                    foreach (var item in data.Dlg.Dlg3)
                    {
                        IRow row = sheet.GetRow(rowIndex3); // Chỉ lấy row, không cần CreateRow
                        if (row != null && item.LocalCode == "V2" && item.IsBold == false)
                        {
                            var a = item.Col1 - data.InputPriceOld.Where(x => x.GoodCode == item.GoodCode).Sum(x => x.FobV2);
                            var b = item.Col2 - data.Dlg.Dlg10.Where(x => x.GoodCode == item.GoodCode).Sum(x => x.Col2);
                            // E11 -> col1
                            ICell cellE = row.GetCell(4);
                            if (cellE != null)
                            {
                                cellE.SetCellValue((double)(a));
                            }

                            // F11 -> col2
                            ICell cellF = row.GetCell(5);
                            if (cellF != null)
                            {
                                cellF.SetCellValue((double)(b));
                            }

                            // G11
                            ICell cellG = row.GetCell(6);
                            if (cellG != null)
                            {
                                cellG.SetCellValue(date);
                            }

                            // G12 -> tangGiam1_2
                            ICell cellH = row.GetCell(7);
                            if (cellH != null)
                            {
                                cellH.SetCellValue((double)(b - a));
                            }

                            ICell cellI = row.GetCell(8);
                            if (cellI != null)
                            {
                                if (a != 0)
                                {
                                    double rateOfIncreaseAndDecrease = (double)((b - a) / a);
                                    cellI.SetCellValue(rateOfIncreaseAndDecrease);
                                }
                                else
                                {
                                    cellH.SetCellValue(0);
                                }
                            }
                            rowIndex3++;
                        }

                    }


                    ISheet sheetCheck = workbook.GetSheetAt(0);
                    IRow rowCheck = sheetCheck.GetRow(4);
                    ICell cellCheck = rowCheck?.GetCell(0);
                    Console.WriteLine($"Giá trị trước khi lưu file: {cellCheck?.StringCellValue}");

                    // Ghi file
                    using (var fsOut = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
                    {
                        workbook.Write(fsOut);
                        Console.WriteLine("Ghi file thành công");
                    }
                    workbook.Close();
                    return $"{folderName}/{fileName}";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi: {ex.Message}");
                throw;
            }
        }
        #endregion

        #region Xuất trình ký Word

        public async Task<string> GenarateWordTrinhKy(string headerId, string nameTemp)
        {
            #region Tạo 1 file word mới từ file template    
            List<string> lstUrl = new List<string>();
            var filePathTemplate = Directory.GetCurrentDirectory() + $"/Template/TempTrinhKy/{nameTemp}.docx";
            var folderName = Path.Combine($"Uploads/Word/{DateTime.Now.ToString("yyyy/MM/dd")}");
            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            var fileName = $"{DateTime.Now.Day}{DateTime.Now.Month}{DateTime.Now.Year}_{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}_{nameTemp}.docx";
            var fullPath = Path.Combine(pathToSave, fileName);
            File.Copy(filePathTemplate, fullPath, true);
            #endregion

            #region Lấy các text element
            List<string> lstTextElement = new List<string>();
            WordDocumentService wordDocumentService = new WordDocumentService();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
            {
                lstTextElement = wordDocumentService.FindTextElement(doc);
                lstTextElement = lstTextElement.Distinct().ToList();
            }
            #endregion

            #region Fill dữ liệu
            var data = await CalculateDiscountOutput(headerId);            
            var header = await _dbContext.TblBuCalculateDiscount.FindAsync(headerId);
            var goods = await _dbContext.TblMdGoods.Where(x=> x.IsActive == true).ToListAsync();

            var goodsList = goods.Where(x => x.IsActive == true)
                                     .OrderBy(x => x.CreateDate)
                                 .ToList();
            var lstSigner = await _dbContext.TblMdSigner.Where(x => x.IsActive == true).ToListAsync();
            var tckt = lstSigner.FirstOrDefault(x => x.Code == header.TcktCode)?.Name ?? "";
            var kdxd = lstSigner.FirstOrDefault(x => x.Code == header.KdxdCode)?.Name ?? "";
            var vietPhuongAn = lstSigner.FirstOrDefault(x => x.Code == header.VietphuonganCode)?.Name ?? "";

            var NguoiKyTen = lstSigner.Where(x => x.Code == header.SignerCode).FirstOrDefault();
            var f_date = $"{header.Date.Day:D2} tháng {header.Date.Month:D2} năm {header.Date.Year}";
            var date = header.Date.ToString("dd/MM/yyyy");
            var hour_now = $"{header.Hour?.Hour:D2} giờ {header.Hour?.Minute:D2} phút";
            var f_date_hour = $"kể từ {header.Date.Hour:D2} giờ 00 ngày {header.Date.Day:D2} tháng {header.Date.Month:D2} năm {header.Date.Year}";

            var calculateDiscountIdOld = await _dbContext.TblBuCalculateDiscount
                        .Where(x => x.Date < header.Date)
                        .Where(x => x.Status == "04")
                        .OrderByDescending(x => x.Date)
                        .FirstOrDefaultAsync();
            var inputPriceOld = new List<TblBuInputPrice>();
                if (calculateDiscountIdOld != null)
            {
                 inputPriceOld = await _dbContext.TblBuInputPrice
                                .Where(x => x.HeaderId == calculateDiscountIdOld.Id).ToListAsync();
            }

            var data_Dlg2_Old = new List<DlgModel>();
            var data_Dlg7_Old = new List<DlgModel>();
            var data_Dlg8_Old = new List<DlgModel>();

            foreach (var g in goodsList)
            {
                foreach (var i in inputPriceOld)
                {
                    data_Dlg2_Old.Add(new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.GblV1 + i.ChenhLech,
                        Col2 = i.GblV2,
                        Col3 = i.GblV2 - i.ChenhLech - i.GblV1,
                    });

                    data_Dlg7_Old.Add(new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = i.L15Blv2,
                        Col4 = Math.Round(i.Vcf * i.L15Blv2),
                        Col5 = Math.Round((((i.Vcf * i.L15Blv2) + i.ThueBvmt) * 1.1M) / 10, 0, MidpointRounding.AwayFromZero) * 10,

                    });

                    data_Dlg8_Old.Add(new DlgModel
                    {
                        GoodCode = i.GoodCode,
                        GoodName = i.GoodName,
                        Col1 = i.Vcf,
                        Col2 = i.ThueBvmt,
                        Col3 = Math.Round(i.ThueBvmt / i.Vcf),
                        Col4 = i.L15Blv2,
                        Col5 = 0,
                        Col6 = Math.Round((i.ThueBvmt / i.Vcf) + i.L15Blv2),
                        Col7 = ((i.ThueBvmt / i.Vcf) + i.L15Blv2) * 1.1M,
                        Note = ""
                    });
                }
            }

            TableCell CreateCell(string text, bool isBold = true, int fontSize = 26, bool isCenter = true, string width = "2400", int? gridSpan = null, int v = 0)
            {
                Run run = new Run(new Text(text));

                RunProperties runProperties = new RunProperties(
                    new RunFonts()
                    {
                        Ascii = "Times New Roman",
                        HighAnsi = "Times New Roman",
                        ComplexScript = "Times New Roman"
                    },
                    new FontSize() { Val = new StringValue(fontSize.ToString()) }
                );

                if (isBold)
                {
                    runProperties.Append(new Bold());
                }

                run.RunProperties = runProperties;

                Paragraph paragraph = new Paragraph(run);

                if (isCenter)
                {
                    paragraph.ParagraphProperties = new ParagraphProperties(
                        new Justification() { Val = JustificationValues.Center }
                    );
                }

                TableCellProperties cellProperties = new TableCellProperties(
                    new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                );

                if (!string.IsNullOrEmpty(width))
                {
                    cellProperties.Append(new TableCellWidth() { Width = width, Type = TableWidthUnitValues.Dxa });
                }

                // Nếu có gridSpan, áp dụng thuộc tính gộp ô
                if (gridSpan.HasValue && gridSpan > 1)
                {
                    cellProperties.Append(new GridSpan() { Val = gridSpan });
                }

                TableCell cell = new TableCell();
                cell.Append(cellProperties);
                cell.Append(paragraph);

                return cell;
            }

            #region fill dữ liệu file
            if (nameTemp == "CongDienKKGiaBanLe")   
            {
                var dlg1 = data.Dlg.Dlg1;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##HOUR_NOW@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour_now);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo ?? "");
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                            case "##TABLE_TT@@":
                                Paragraph paragraph = body.Descendants<Paragraph>()
                                               .FirstOrDefault(p => p.InnerText.Contains("##TABLE_TT@@"));
                                if (paragraph != null)
                                {
                                    Table table = new Table();
                                    DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                                       new TableCellMarginDefault(
                                           new LeftMargin() { Width = "115" },
                                           new RightMargin() { Width = "115" },
                                           new TopMargin() { Width = "50" },
                                           new BottomMargin() { Width = "50" }

                                       )
                                   );
                                    table.AppendChild(tblProperties);

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in goodsList)
                                    {
                                        var item = data.Dlg.Dlg2.FirstOrDefault(x => x.GoodCode == i.Code);
                                        var itemOld = data_Dlg2_Old.FirstOrDefault(x => x.GoodCode == i.Code);
                                        if (item != null)
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.Name, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(item.Col1.ToString("N0"), true, 30, false, "2000"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                            row.Append(CreateCell(calculateDiscountIdOld == null || itemOld?.Col1 != item.Col1 ? "(Thay đổi)" : "(Không thay đổi)", false, 26, false, "2800"));
                                            table.Append(row);
                                            o++;
                                        }
                                    }
                                    #endregion
                                    paragraph.Parent.InsertAfter(table, paragraph);
                                    paragraph.Remove();
                                }
                                break;
                            case "##TABLE_VCL@@":
                                Paragraph paragraph2 = body.Descendants<Paragraph>()
                                           .FirstOrDefault(p => p.InnerText.Contains("##TABLE_VCL@@"));
                                if (paragraph2 != null)
                                {
                                    Table table = new Table();
                                    DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                                       new TableCellMarginDefault(
                                           new LeftMargin() { Width = "115" },
                                           new RightMargin() { Width = "115" },
                                           new TopMargin() { Width = "50" },
                                           new BottomMargin() { Width = "50" }

                                       )
                                   );
                                    table.AppendChild(tblProperties);

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in data.Dlg.Dlg2)
                                    {
                                            var itemOld = data_Dlg2_Old.FirstOrDefault(x => x.GoodCode == i.GoodCode);
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.GoodName, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(i.Col2.ToString("N0"), true, 30, false, "2000"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                            row.Append(CreateCell(calculateDiscountIdOld == null || itemOld?.Col2 != i.Col2 ? "(Thay đổi)" : "(Không thay đổi)", false, 26, false, "2800"));
                                            table.Append(row);
                                            o++;
                                    }

                                    #endregion

                                    paragraph2.Parent.InsertAfter(table, paragraph2);
                                    paragraph2.Remove();
                                }
                                break;
                        }
                    }
                }
                //}

                //if (code != lstCustomerChecked.LastOrDefault())
                //{
                //    AppendWordFilesToNewDocument(filePathTemplate, fullPath);
                //}
            }

            else if (nameTemp == "QDGNoiDung")
            {
                var dlg7 = data.Dlg.Dlg7;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##HOUR_NOW@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour_now);
                                break;
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo);
                                break;
                            case "##CONG_DIEN_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.CongDienSo ?? "");
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                            case "##TABLE_ND@@":
                                Paragraph paragraph = body.Descendants<Paragraph>()
                                           .FirstOrDefault(p => p.InnerText.Contains("##TABLE_ND@@"));
                                if (paragraph != null)
                                {
                                    Table table = new Table();
                                    TableProperties tblProperties = new TableProperties(
                                        new TableBorders(
                                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                                        ),
                                        new TableCellMarginDefault(
                                            new LeftMargin() { Width = "115" },
                                            new RightMargin() { Width = "115" },
                                            new TopMargin() { Width = "50" },
                                            new BottomMargin() { Width = "50" }
                                        )
                                    );
                                    table.AppendChild(tblProperties);

                                    #region Header table
                                    TableRow rowHeader = new TableRow();

                                    TableCell cell1 = CreateCell("TT", true, 26, true, "1");
                                    TableCell cell2 = CreateCell("MẶT HÀNG", true, 26, true, "3000");
                                    TableCell cell3 = CreateCell("ĐƠN GIÁ", true, 26, true, "3000");
                                    TableCell cell4 = CreateCell("ĐƠN VỊ TÍNH", true, 26, true, "3000");
                                    TableCell cell5 = CreateCell("GHI CHÚ", true, 26, true, "3000");

                                    rowHeader.Append(cell1);
                                    rowHeader.Append(cell2);
                                    rowHeader.Append(cell3);
                                    rowHeader.Append(cell4);
                                    rowHeader.Append(cell5);

                                    table.Append(rowHeader);
                                    #endregion

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in dlg7)
                                    {
                                            var itemOld = data_Dlg7_Old.FirstOrDefault(x => x.GoodCode == i.GoodCode);

                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.Stt, true, 26, true, "1"));
                                            row.Append(CreateCell(i.GoodName, true, 26, true));
                                            row.Append(CreateCell(i.Col5.ToString("N0"), true, 30));
                                            row.Append(CreateCell("đ/ lít thực tế", true, 26));
                                            row.Append(CreateCell(calculateDiscountIdOld == null || itemOld?.Col5 != i.Col5 ? "Thay đổi" : "Không thay đổi", true, 26));
                                            table.Append(row);
                                            o++;
                                    }
                                    #endregion

                                    paragraph.Parent.InsertAfter(table, paragraph);
                                    paragraph.Remove();
                                }
                                break;

                        }

                    }
                }
            }

            else if (nameTemp == "QDGCtyPTS")
            {
                var dlg8 = data.Dlg.Dlg8;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##HOUR_NOW@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour_now);
                                break;
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo);
                                break;
                            case "##CONG_DIEN_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.CongDienSo ?? "");
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                            case "##TABLE_PTS@@":
                                Paragraph paragraph = body.Descendants<Paragraph>()
                                           .FirstOrDefault(p => p.InnerText.Contains("##TABLE_PTS@@"));
                                if (paragraph != null)
                                {
                                    Table table = new Table();
                                    TableProperties tblProperties = new TableProperties(
                                        new TableBorders(
                                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                                        ),
                                        new TableCellMarginDefault(
                                            new LeftMargin() { Width = "115" },
                                            new RightMargin() { Width = "115" },
                                            new TopMargin() { Width = "50" },
                                            new BottomMargin() { Width = "50" }
                                        )
                                    );
                                    table.AppendChild(tblProperties);

                                    #region Header table
                                    TableRow rowHeader = new TableRow();

                                    TableCell cell1 = CreateCell("TT", true, 26, true, "1");
                                    TableCell cell2 = CreateCell("MẶT HÀNG", true, 26, true, "3000");
                                    TableCell cell3 = CreateCell("GIÁ BÁN", true, 26, true, "3000");
                                    TableCell cell4 = CreateCell("ĐƠN VỊ TÍNH", true, 26, true, "3000");
                                    TableCell cell5 = CreateCell("GHI CHÚ", true, 26, true, "3000");

                                    rowHeader.Append(cell1);
                                    rowHeader.Append(cell2);
                                    rowHeader.Append(cell3);
                                    rowHeader.Append(cell4);
                                    rowHeader.Append(cell5);

                                    table.Append(rowHeader);
                                    #endregion

                                    #region Gendata table
                                    foreach (var i in dlg8)
                                    {
                                            var itemOld = data_Dlg8_Old.FirstOrDefault(x => x.GoodCode == i.GoodCode);
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.Stt, true, 26, true, "1"));
                                            row.Append(CreateCell(i.GoodName, true, 26, true, "3000"));
                                            row.Append(CreateCell(i.Col6.ToString("N0"), true, 30, true, "3000"));
                                            row.Append(CreateCell("đ/ lít 15", true, 26, true, "3000"));
                                            row.Append(CreateCell(calculateDiscountIdOld == null || itemOld?.Col6 != i.Col6 ? "Thay đổi" : "Không thay đổi", true, 26));
                                            table.Append(row);
                                    }
                                    #endregion

                                    paragraph.Parent.InsertAfter(table, paragraph);
                                    paragraph.Remove();
                                }
                                break;
                        }

                    }
                }
            }

            else if (nameTemp == "QDGBanBuon")
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##F_DATE_HOUR@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date_hour);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo);
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                        }

                    }
                }
            }

            else if (nameTemp == "MucGiamGiaNQTM")
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##F_DATE_HOUR@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date_hour);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo);
                                break;

                            case "##KDXD@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, kdxd);
                                break;

                            case "##TCKT@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, tckt);
                                break;

                            case "##VIET_PHUONG_AN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, vietPhuongAn);
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                        }

                    }
                }
            }

            else if (nameTemp == "KeKhaiGia")
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##HOUR@@":
                                var hour = $"{header.Date.Hour:D2}";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour);
                                break;
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                        }

                    }
                }
            }

            else if (nameTemp == "ToTrinh")
            {
                var config = new MapperConfiguration(cfg =>
                {
                    cfg.CreateMap<TblMdGoods, GoodsDto>(); // Thêm các map cần thiết
                });
                var mapper = config.CreateMapper();

                var DIService = new DiscountInformationService(_dbContext, mapper);
                var lstDI = await DIService.getAll(headerId);
                //var lstDI = await _discountService.getAll(headerId);
                var dlg9 = data.Dlg.Dlg9;
                var dlg10 = data.Dlg.Dlg10;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;
                    if (lstDI.discount.Count() == 0)
                    {
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLX95@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLXE5@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLX05@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLX01@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##AP95@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##APE5@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##AP05@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##AP01@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PV95@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PVE5@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PV05@@", "");
                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PV01@@", "");

                    }
                    else
                    {

                        foreach (var i in lstDI.discount[0]?.CK)
                        {
                            if (i.GoodsCode == "0201032")
                            {
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLX95@@", i.plxna.ToString());

                                foreach (var a in i.DT)
                                {
                                    if(a.code == "APP")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##AP95@@", a.ckCl[0].ToString() ?? "0");
                                    }
                                    if (a.code == "PVOIL")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PV95@@", a.ckCl[0].ToString() ?? "0");

                                    }
                                }
                            }
                            if (i.GoodsCode == "0201004")
                            {
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLXE5@@", i.plxna.ToString());

                                foreach (var a in i.DT)
                                {
                                    if (a.code == "APP")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##APE5@@", a.ckCl[0].ToString() ?? "0");
                                    }
                                    if (a.code == "PVOIL")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PVE5@@", a.ckCl[0].ToString() ?? "0");

                                    }
                                }
                            }
                            if (i.GoodsCode == "0601005")
                            {
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLX01@@", i.plxna.ToString());

                                foreach (var a in i.DT)
                                {
                                    if (a.code == "APP")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##AP01@@", a.ckCl[0].ToString() ?? "0");
                                    }
                                    if (a.code == "PVOIL")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PV01@@", a.ckCl[0].ToString() ?? "0");

                                    }
                                }
                            }
                            if (i.GoodsCode == "0601002")
                            {
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PLX05@@", i.plxna.ToString());

                                foreach (var a in i.DT)
                                {
                                    if (a.code == "APP")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##AP05@@", a.ckCl[0].ToString() ?? "0");
                                    }
                                    if (a.code == "PVOIL")
                                    {
                                        wordDocumentService.ReplaceStringInWordDocumennt(doc, "##PV05@@", a.ckCl[0].ToString() ?? "0");

                                    }
                                }
                            }
                        }

                    }

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##HOUR@@":
                                var hour = $"{header.Date.Hour:D2}";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour);
                                break;
                            case "##TABLE_LAI_GOP@@":
                                Paragraph paragraph = body.Descendants<Paragraph>()
                                       .FirstOrDefault(p => p.InnerText.Contains("##TABLE_LAI_GOP@@"));
                                if (paragraph != null)
                                {
                                    Table table = new Table();
                                    TableProperties tblProperties = new TableProperties(
                                        new TableBorders(
                                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                                        )
                                    );
                                    table.AppendChild(tblProperties);

                                    #region Header table


                                    TableRow headerRow1 = new TableRow();
                                    headerRow1.Append(CreateCell("So sánh lãi gộp giữa hai kỳ", true, 26, true, "2082", 4, 1));

                                    TableRow headerRow2 = new TableRow();
                                    headerRow2.Append(CreateCell("Mặt hàng", true, 26, true, "2082"));
                                    headerRow2.Append(CreateCell("LG Cũ", true, 26, true, "2082"));
                                    headerRow2.Append(CreateCell("LG Mới", true, 26, true, "2082"));
                                    headerRow2.Append(CreateCell("Tăng/Giảm", true, 26, true, "2082"));

                                    table.Append(headerRow1);
                                    table.Append(headerRow2);

                                    // Thêm dòng tiêu đề "Vùng thị trường trung tâm"
                                    TableRow regionRow1 = new TableRow();
                                    regionRow1.Append(CreateCell("Vùng thị trường trung tâm", true, 26, true, "2082", 4, 1));
                                    table.Append(regionRow1);
                                    #endregion

                                    #region Gendata table                                   
                                    foreach (var i in dlg9.Where(x => x.LocalCode == "V1"))
                                    {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.GoodName, false, 26, false, "3500")); // Tên mặt hàng
                                            row.Append(CreateCell(i.Col1.ToString("N0"), false, 26, false, "2082")); // LG cũ
                                            row.Append(CreateCell(i.Col2.ToString("N0"), false, 26, false, "2082")); // LG mới
                                            row.Append(CreateCell(i.Col3.ToString("N0"), false, 26, false, "2082"));
                                            table.Append(row);
                                    }

                                    // Thêm dòng tiêu đề "Các vùng thị trường còn lại"
                                    TableRow regionRow2 = new TableRow();
                                    regionRow2.Append(CreateCell("Các vùng thị trường còn lại", true, 26, true, "2082", 4, 1)); // Gộp 4 cột
                                    table.Append(regionRow2);

                                    // Duyệt danh sách dlg6, in từng mặt hàng thuộc vùng còn lại
                                    foreach (var i in dlg9.Where(x => x.LocalCode == "V2"))
                                    {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.GoodName, false, 26, false, "3500")); // Tên mặt hàng
                                            row.Append(CreateCell(i.Col1.ToString("N0"), false, 26, false, "2082")); // LG cũ
                                            row.Append(CreateCell(i.Col2.ToString("N0"), false, 26, false, "2082")); // LG mới
                                            row.Append(CreateCell(i.Col3.ToString("N0"), false, 26, false, "2082"));
                                            table.Append(row);
                                    }
                                    #endregion

                                    paragraph.Parent.InsertAfter(table, paragraph);
                                    paragraph.Remove();
                                }
                                break;
                            case "##TABLE_GIAM_GIA@@":
                                Paragraph paragraph1 = body.Descendants<Paragraph>()
                                        .FirstOrDefault(p => p.InnerText.Contains("##TABLE_GIAM_GIA@@"));
                                if (paragraph1 != null)
                                {
                                    Table table = new Table();
                                    TableProperties tblProperties = new TableProperties(
                                        new TableBorders(
                                            new TopBorder { Val = BorderValues.Single, Size = 4 },
                                            new BottomBorder { Val = BorderValues.Single, Size = 4 },
                                            new LeftBorder { Val = BorderValues.Single, Size = 4 },
                                            new RightBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                                            new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                                        )
                                    );
                                    table.AppendChild(tblProperties);

                                    #region Header table
                                    TableRow headerRow1 = new TableRow();

                                    headerRow1.Append(CreateCell("So sánh lãi gộp giữa hai kỳ", true, 26, true, "2082", 4, 1));

                                    TableRow headerRow2 = new TableRow();
                                    headerRow2.Append(CreateCell("Mặt hàng", true, 26, true, "2082"));
                                    headerRow2.Append(CreateCell("CK Cũ", true, 26, true, "2082"));
                                    headerRow2.Append(CreateCell("CK Mới", true, 26, true, "2082"));
                                    headerRow2.Append(CreateCell("Tăng/Giảm", true, 26, true, "2082"));

                                    table.Append(headerRow1);
                                    table.Append(headerRow2);

                                    #endregion

                                    #region Gendata table
                                    foreach (var i in dlg10)
                                    {
                                            TableRow row = new TableRow();

                                            row.Append(CreateCell(i.GoodName, false, 26, false, "3500")); // Tên mặt hàng
                                            row.Append(CreateCell(i.Col1.ToString("N0"), false, 26, false, "2082")); // LG cũ
                                            row.Append(CreateCell(i.Col2.ToString("N0"), false, 26, false, "2082")); // LG mới
                                            row.Append(CreateCell(i.Col3.ToString("N0"), false, 26, false, "2082"));
                                            table.Append(row);

                                    }
                                    #endregion

                                    paragraph1.Parent.InsertAfter(table, paragraph1);
                                    paragraph1.Remove();
                                }
                                break;
                        }

                    }
                }
            }

            else if (nameTemp == "QDGBanLe")
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##HOUR_NOW@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour_now);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo);
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                            case "##TABLE_TT@@":
                                Paragraph paragraph = body.Descendants<Paragraph>()
                                           .FirstOrDefault(p => p.InnerText.Contains("##TABLE_TT@@"));
                                if (paragraph != null)
                                {
                                    Table table = new Table();
                                    DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                                       new TableCellMarginDefault(
                                           new LeftMargin() { Width = "115" },
                                           new RightMargin() { Width = "115" },
                                           new TopMargin() { Width = "50" },
                                           new BottomMargin() { Width = "50" }

                                       )
                                   );
                                    table.AppendChild(tblProperties);

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in goodsList)
                                    {

                                        var item = data.Dlg.Dlg2.FirstOrDefault(x => x.GoodCode == i.Code);
                                        var itemOld = data_Dlg2_Old.FirstOrDefault(x => x.GoodCode == i.Code);
                                        if (item != null)
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.Name, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(item.Col1.ToString("N0"), true, 26, false, "2400"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                            row.Append(CreateCell(calculateDiscountIdOld == null || itemOld?.Col1 != item.Col1 ? "(Thay đổi)" : "(Không thay đổi)", false, 26, false, "2900"));
                                            table.Append(row);
                                            o++;
                                        }
                                    }
                                    #endregion
                                    paragraph.Parent.InsertAfter(table, paragraph);
                                    paragraph.Remove();
                                }
                                break;
                            case "##TABLE_CL@@":
                                Paragraph paragraph2 = body.Descendants<Paragraph>()
                                           .FirstOrDefault(p => p.InnerText.Contains("##TABLE_CL@@"));
                                if (paragraph2 != null)
                                {
                                    Table table = new Table();
                                    DocumentFormat.OpenXml.Wordprocessing.TableProperties tblProperties = new DocumentFormat.OpenXml.Wordprocessing.TableProperties(
                                       new TableCellMarginDefault(
                                           new LeftMargin() { Width = "115" },
                                           new RightMargin() { Width = "115" },
                                           new TopMargin() { Width = "50" },
                                           new BottomMargin() { Width = "50" }

                                       )
                                   );
                                    table.AppendChild(tblProperties);

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in goodsList)
                                    {
                                        var item = data.Dlg.Dlg2.FirstOrDefault(x => x.GoodCode == i.Code);
                                        var itemOld = data_Dlg2_Old.FirstOrDefault(x => x.GoodCode == i.Code);
                                        TableRow row = new TableRow();
                                        row.Append(CreateCell("+ " + i.Name, true, 26, false, "3500"));
                                        row.Append(CreateCell(":", false, 26, true, "1"));
                                        row.Append(CreateCell(item.Col2.ToString("N0"), true, 26, false, "2400"));
                                        row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                        row.Append(CreateCell(calculateDiscountIdOld == null || itemOld?.Col2 != item.Col2 ? "(Thay đổi)" : "(Không thay đổi)", false, 26, false, "2900"));
                                        table.Append(row);
                                        o++;
                                    }

                                   #endregion

                                    paragraph2.Parent.InsertAfter(table, paragraph2);
                                    paragraph2.Remove();
                                }
                                break;
                        }

                    }
                }
            }
            
            else if (nameTemp == "ThueBvmtPts")
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;
                    foreach( var i in data.Dlg.Dlg4)
                    {
                        if (i.GoodCode == "0201032")
                        {
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##VCF95@@", (i.Col1).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##BV_95@@", (i.Col2).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##L15_95@@", (Math.Round(i.Col2 / i.Col1, 0)).ToString());
                        }
                        else if (i.GoodCode == "0201004")
                        {
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##VCF92@@", (i.Col1).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##BV_92@@", (i.Col2).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##L15_92@@", (Math.Round(i.Col2 / i.Col1, 0)).ToString());
                        }
                        else if (i.GoodCode == "0601005")
                        {
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##VCF01@@", (i.Col1).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##BV_01@@", (i.Col2).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##L15_01@@", (Math.Round(i.Col2 / i.Col1, 0)).ToString());
                        }
                        else if (i.GoodCode == "0601002")
                        {
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##VCF05@@", (i.Col1).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##BV_05@@", (i.Col2).ToString());
                            wordDocumentService.ReplaceStringInWordDocumennt(doc, "##L15_05@@", (Math.Round(i.Col2 / i.Col1, 0)).ToString());
                        }
                    }


                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##F_DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date);
                                break;
                            case "##DATE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, date);
                                break;
                            case "##HOUR_NOW@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, hour_now);
                                break;
                            case "##CONG_DIEN_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.CongDienSo ?? "");
                                break;
                            case "##VAN_BAN_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.VanBanSo ?? "");
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, NguoiKyTen.Name);
                                break;
                            case "##KDXD@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, kdxd);
                                break;

                            case "##TCKT@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, tckt);
                                break;
                            case "##VIET_PHUONG_AN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, vietPhuongAn);
                                break;

                        }

                    }
                }
            }
            #endregion


            #endregion

            return $"{folderName}/{fileName}";
        }
        #endregion

        #region Xuất trình ký

        static void AppendWordFilesToNewDocument(string directoryPath, string newWordFilePath)
        {
            using (WordprocessingDocument sourceDocument = WordprocessingDocument.Open(directoryPath, false))
            {
                DocumentFormat.OpenXml.Wordprocessing.Body sourceBody = sourceDocument.MainDocumentPart.Document.Body;

                using (WordprocessingDocument destinationDocument = WordprocessingDocument.Open(newWordFilePath, true))
                {
                    DocumentFormat.OpenXml.Wordprocessing.Body destinationBody = destinationDocument.MainDocumentPart.Document.Body;

                    // Đảm bảo có ngắt trang trước khi chèn nội dung mới
                    var lastParagraph = destinationBody.Elements<Paragraph>().LastOrDefault();
                    if (lastParagraph == null || !lastParagraph.InnerText.Contains("\f"))
                    {
                        destinationBody.AppendChild(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
                    }

                    // Chèn nội dung từ file gốc
                    foreach (var element in sourceBody.Elements())
                    {
                        destinationBody.Append(element.CloneNode(true));
                    }

                    destinationDocument.MainDocumentPart.Document.Save();
                }
            }
        }


        public async Task<string> GenarateWord(List<CustomBBDOExportWord> lstCustomerChecked, string headerId)
        {
            try
            {
                var s = new ExportWordService(_dbContext);
                var data = await CalculateDiscountOutput(headerId);
                return await s.GenarateWord(lstCustomerChecked, headerId, data);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<string> GenarateWord2(List<CustomBBDOExportWord> lstCustomerChecked, string headerId)
        {
            #region Tạo 1 file word mới từ file template
            var filePathTemplate = Directory.GetCurrentDirectory() + "/Template/ThongBaoGia.docx";
            var folderName = Path.Combine($"Uploads/Word/{DateTime.Now.ToString("yyyy/MM/dd")}");
            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            var fileName = $"{DateTime.Now.Day}{DateTime.Now.Month}{DateTime.Now.Year}_{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}_ThongBaoGia.docx";
            var fullPath = Path.Combine(pathToSave, fileName);
            File.Copy(filePathTemplate, fullPath, true);
            #endregion

            #region Lấy các text element
            List<string> lstTextElement = new List<string>();
            WordDocumentService wordDocumentService = new WordDocumentService();
            using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
            {
                lstTextElement = wordDocumentService.FindTextElement(doc);
                lstTextElement = lstTextElement.Distinct().ToList();
            }
            #endregion

            #region Fill dữ liệu 
            var data = await CalculateDiscountOutput(headerId);
            var header = await _dbContext.TblBuCalculateDiscount.FindAsync(headerId);
            var sinner = await _dbContext.TblMdSigner.FirstOrDefaultAsync(x => x.Code == header.SignerCode);
            var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).ToListAsync();
            var IsN1 = false;
            foreach (var l in lstCustomerChecked)
            {
                var d = data.Vk11Bb.Where(x => x.Col4 == l.code).ToList();
                var c = await _dbContext.TblBuInputCustomerBbdo.FirstOrDefaultAsync(x => x.Code == l.code);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##DATE@@":
                                var text = $"{header?.Date.Hour:D2}h00 ngày {header?.Date.Day} tháng {header?.Date.Month} năm {header?.Date.Year}";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, text);
                                break;
                            case "##DATE2@@":
                                var text2 = $"ngày {header?.Date.Day} tháng {header?.Date.Month} năm {header?.Date.Year}";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, text2);
                                break;
                            case "##QUYET_DINH_SO@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, header.QuyetDinhSo ?? "");
                                break;
                            case "##DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, sinner.Code != "TongGiamDoc" ? "KT.GIÁM ĐỐC CÔNG TY" : "");
                                break;
                            case "##NGUOI_DAI_DIEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, sinner.Position);
                                break;
                            case "##TEN@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, sinner.Name);
                                break;
                            case "##COMPANY@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, c?.Name);
                                break;
                            case "##ADDRESS@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, c?.Adrress);
                                break;
                            case "##TABLE@@":
                                
                                break;
                            case "##CHIPHI@@":
                                var chiphi = IsN1 == false ? "và chi phí vân chuyển" : "";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, chiphi);
                                break;
                        }
                    }
                }

                if (l.code != lstCustomerChecked.LastOrDefault().code)
                {
                    AppendWordFilesToNewDocument(filePathTemplate, fullPath);
                }
            }
            #endregion

            return $"{folderName}/{fileName}";
        }

        public async Task<List<TblBuInputCustomerBbdo>> GetCustomerBbdo(string id)
        {
            try
            {
                var query = await _dbContext.TblBuInputCustomerBbdo
                    .Where(x => x.HeaderId == id)
                    .GroupBy(x => x.Code)
                    .Select(g => g.FirstOrDefault())
                    .ToListAsync();
                return query;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }

        public async Task<List<CustomInput>> GetAllInputCustomer()
        {
            try
            {
                var customerBbdo = await _dbContext.TblMdCustomerBbdo
                    .GroupBy(x => x.Code)
                    .Select(g => g.Select(x => new CustomInput { code = x.Code, name = x.Name }).FirstOrDefault())
                    .ToListAsync();

                var customerPt = await _dbContext.TblMdCustomerPt
                    .GroupBy(x => x.Code)
                    .Select(g => g.Select(x => new CustomInput { code = x.Code, name = x.Name }).FirstOrDefault())
                    .ToListAsync();

                var customerFob = await _dbContext.TblMdCustomerFob
                    .GroupBy(x => x.Code)
                    .Select(g => g.Select(x => new CustomInput { code = x.Code, name = x.Name }).FirstOrDefault())
                    .ToListAsync();

                var customerDb = await _dbContext.TblMdCustomerDb
                    .GroupBy(x => x.Code)
                    .Select(g => g.Select(x => new CustomInput { code = x.Code, name = x.Name }).FirstOrDefault())
                    .ToListAsync();

                var customerPts = await _dbContext.TblMdCustomerPts
                    .GroupBy(x => x.Code)
                    .Select(g => g.Select(x => new CustomInput { code = x.Code, name = x.Name }).FirstOrDefault())
                    .ToListAsync();

                var customerTnpp = await _dbContext.TblMdCustomerTnpp
                    .GroupBy(x => x.Code)
                    .Select(g => g.Select(x => new CustomInput { code = x.Code, name = x.Name }).FirstOrDefault())
                    .ToListAsync();

                // Gộp danh sách nhưng đảm bảo kiểu trả về là List<CustomInput>
                var result = customerBbdo
                    .Concat(customerPt)
                    .Concat(customerFob)
                    .Concat(customerDb)
                    .Concat(customerPts)
                    .Concat(customerTnpp)
                    .ToList(); // Loại bỏ kiểu dynamic

                return result;
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return new List<CustomInput>(); // Trả về danh sách rỗng thay vì null
            }
        }


        public async Task<List<string>> GenarateFile(List<string> lstCustomerChecked, string type, string headerId, CalculateDiscountInputModel data, List<CustomBBDOExportWord>? lstCustomerCheckedWord = null)
        {

            if (type == "WORD")
            {
                var path = await GenarateWord(lstCustomerCheckedWord, headerId);
                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = path.Replace($"Uploads/Word/{DateTime.Now.ToString("yyyy/MM/dd")}/", ""),
                    Type = "docx",
                    Path = path
                });
                await _dbContext.SaveChangesAsync();
                return new List<string> { path };
            }
            if (type == "WORDTRINHKY")
            {
                List<string> lstUrl = new List<string>();
                foreach (var n in lstCustomerChecked)
                {
                    if (n == "KeKhaiGiaChiTiet")
                    {
                        var path = await ExportExcelTrinhKy(headerId);
                        _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                        {
                            Code = Guid.NewGuid().ToString(),
                            HeaderCode = headerId,
                            Name = path.Replace($"Uploads/Excel/{DateTime.Now.ToString("yyyy/MM/dd")}/", ""),
                            Type = "xlsx",
                            Path = path
                        });
                        lstUrl.Add(path);
                        await _dbContext.SaveChangesAsync();
                    }
                    else
                    {
                        var path = await GenarateWordTrinhKy(headerId, n);
                        _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                        {
                            Code = Guid.NewGuid().ToString(),
                            HeaderCode = headerId,
                            Name = path.Replace($"Uploads/Word/{DateTime.Now.ToString("yyyy/MM/dd")}/", ""),
                            Type = "docx",
                            Path = path
                        });
                        await _dbContext.SaveChangesAsync();
                        lstUrl.Add(path);

                    }
                }
                return lstUrl;
            }
            else
            {
                var w = await GenarateWord(lstCustomerCheckedWord, headerId);
                var pathWord = Directory.GetCurrentDirectory() + "/" + w;
                Aspose.Words.Document doc = new Aspose.Words.Document(pathWord);
                var folderName = Path.Combine($"Uploads/Pdf/{DateTime.Now.ToString("yyyy/MM/dd")}");
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                var fileName = $"ThongBaoGia_{DateTime.Now:ddMMyyyy_HHmmss}.pdf";
                var fullPath = Path.Combine(Directory.GetCurrentDirectory(), folderName, fileName);
                doc.Save(fullPath, SaveFormat.Pdf);

                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = fileName,
                    Type = "pdf",
                    Path = $"{folderName}/{fileName}",
                });
                await _dbContext.SaveChangesAsync();
                return new List<string> { $"{folderName}/{fileName}" };
                //return ;
            }
        }
        public async Task  GenarateFileMail(List<string> lstCustomerChecked, string type, string headerId, List<CustomBBDOExportWord>? lstCustomerCheckedWord = null)
        {

            var s = new ExportWordService(_dbContext);
            var data = await CalculateDiscountOutput(headerId);

            foreach ( var item in lstCustomerCheckedWord)
            {
                var listparam = new List<CustomBBDOExportWord>();
                listparam.Add(item);

                var path = await s.GenarateWord(listparam, headerId, data);

                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = path.Replace($"Uploads/Word/{DateTime.Now.ToString("yyyy/MM/dd")}/", ""),
                    Type = "docx",
                    Path = path,
                    CustomerCode = item.code
                });
            }
            await _dbContext.SaveChangesAsync();
        }

        public async Task<string> SaveSMS(string headerId, string smsName)
        {

            if (smsName == "SMS thông báo thù lao")
            {
                var lstSMS = _dbContext.TblCmNotifySms.FirstOrDefault(x => x.HeaderId == headerId && x.Status == "TBTL");

                if (lstSMS != null)
                {
                    return "01";
                }
            }else if (smsName == "SMS Thông báo giá bán lẻ niêm yết")
            {
                var lstSMS = _dbContext.TblCmNotifySms.FirstOrDefault(x => x.HeaderId == headerId && x.Status == "TBGBL");

                if (lstSMS != null)
                {
                    return "02";
                }
            }
            var data = await this.CalculateDiscountOutput(headerId);
            var dataHeader = await this.GetInput(headerId);

            var Template = _dbContext.TblAdConfigTemplate.Where(x => x.Name == smsName).FirstOrDefault();
            try
            {
                DateTime Date = dataHeader.Header.Date;
                var Ngay = $"{Date.Hour:D2}h ngày {Date:dd/MM/yyyy}";

                if (smsName == "SMS thông báo thù lao")
                {
                    foreach (var i in data.Pl1)
                    {
                        if (i.MarketCode != null)
                        {
                            var lstPhonecustomer = _dbContext.TblMdCustomerPhone.Where(x => x.MarketCode == i.MarketCode).Where(x => x.IsActive == true).ToList();

                            foreach (var p in lstPhonecustomer)
                            {
                                var goods = $"X 95-III: {i.Col1.ToString("N0")} d/l, X E5-II {i.Col2.ToString("N0")} d/l, Do 0,001S: {i.Col3.ToString("N0")} d/l, Do 0,05S-II: {i.Col4.ToString("N0")} d/l";

                                var info = new TblNotifySms()
                                {
                                    Id = Guid.NewGuid().ToString(),
                                    PhoneNumber = p.Phone,
                                    Subject = Template.Title ?? "",
                                    Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[market]", i.MarketName).Replace("[goods]", goods),
                                    IsSend = "C",
                                    Status = "TBTL",
                                    CustomerCode = "_",
                                    MarketCode = i.MarketCode,
                                    NumberRetry = 0,
                                    HeaderId = headerId
                                };
                                _dbContext.TblCmNotifySms.Add(info);
                            }
                        }
                    }

                    var pl2 = data.Pl2;

                    foreach (var i in pl2)
                    {
                        if (i.CustomerCode != null)
                        {
                            var customer = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).FirstOrDefault();

                            if (customer != null)
                            {
                                var market = _dbContext.TblMdMarket.Where(x => x.Code == i.MarketCode).Select(x => x.Name).FirstOrDefault();

                                var goods = $"X 95-III: {i.Col1.ToString("N0")} d/l, X E5-II {i.Col2.ToString("N0")} d/l, Do 0,001S: {i.Col3.ToString("N0")} d/l, Do 0,05S-II: {i.Col4.ToString("N0")} d/l";

                                var info = new TblNotifySms()
                                {
                                    Id = Guid.NewGuid().ToString(),
                                    PhoneNumber = customer.Phone,
                                    Subject = Template.Title ?? "",
                                    Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[market]", market).Replace("[goods]", goods),
                                    IsSend = "C",
                                    Status = "TBTL",
                                    CustomerCode = i.CustomerCode,
                                    MarketCode = i.MarketCode,
                                    NumberRetry = 0,
                                    HeaderId = headerId
                                };
                                _dbContext.TblCmNotifySms.Add(info);
                            }
                        }
                    }


                    var pl3 = data.Pl3;

                    foreach (var i in pl3)
                    {
                        var customer = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).FirstOrDefault();

                        if (customer != null)
                        {
                            var market = "kho N.Huong/Ben Thuy";

                            var goods = $"X 95-III: {i.Col1.ToString("N0")} d/l, X E5-II {i.Col2.ToString("N0")} d/l, Do 0,001S: {i.Col3.ToString("N0")} d/l, Do 0,05S-II: {i.Col4.ToString("N0")} d/l";

                            var info = new TblNotifySms()
                            {
                                Id = Guid.NewGuid().ToString(),
                                PhoneNumber = customer.Phone,
                                Subject = Template.Title ?? "",
                                Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[market]", market).Replace("[goods]", goods),
                                IsSend = "C",
                                Status = "TBTL",
                                CustomerCode = i.CustomerCode,
                                MarketCode = i.MarketCode,
                                NumberRetry = 0,
                                HeaderId = headerId
                            };
                            _dbContext.TblCmNotifySms.Add(info);

                        }
                    }


                    var pl4 = data.Pl4;

                    foreach (var i in pl4)
                    {
                        var customer = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).FirstOrDefault();

                        if (customer != null)
                        {
                            var market = "kho N.Huong/Ben Thuy";


                            var goods = $"X 95-III: {i.Col1.ToString("N0")} d/l, X E5-II {i.Col2.ToString("N0")} d/l, Do 0,001S: {i.Col3.ToString("N0")} d/l, Do 0,05S-II: {i.Col4.ToString("N0")} d/l";

                            var info = new TblNotifySms()
                            {
                                Id = Guid.NewGuid().ToString(),
                                PhoneNumber = customer.Phone,
                                Subject = Template.Title ?? "",
                                Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[market]", market).Replace("[goods]", goods),
                                IsSend = "C",
                                Status = "TBTL",
                                CustomerCode = i.CustomerCode,
                                MarketCode = i.MarketCode,
                                NumberRetry = 0,
                                HeaderId = headerId
                            };
                            _dbContext.TblCmNotifySms.Add(info);

                        }
                    }
                }
                else if (smsName == "SMS Thông báo giá bán lẻ niêm yết")
                {
                    var dlgV1 = data.Dlg.Dlg3.Where(x => x.LocalCode == "V1").ToList();
                    var goodsV1 = $"X 95-III: {dlgV1.Where(x => x.GoodCode == "0201032").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV1.Where(x => x.GoodCode == "0201032").Sum(x => x.Col3).ToString("N0")} d/l), " +
                        $"X E5-II: {dlgV1.Where(x => x.GoodCode == "0201004").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV1.Where(x => x.GoodCode == "0201004").Sum(x => x.Col3).ToString("N0")} d/l), " +
                        $"Do 0,001S: {dlgV1.Where(x => x.GoodCode == "0601005").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV1.Where(x => x.GoodCode == "0601005").Sum(x => x.Col3).ToString("N0")} d/l), " +
                        $"Do 0,05S-II: {dlgV1.Where(x => x.GoodCode == "0601002").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV1.Where(x => x.GoodCode == "0601002").Sum(x => x.Col3).ToString("N0")} d/l)";

                    var dlgV2 = data.Dlg.Dlg3.Where(x => x.LocalCode == "V2").ToList();
                    var goodsV2 = $"X 95-III: {dlgV2.Where(x => x.GoodCode == "0201032").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV2.Where(x => x.GoodCode == "0201032").Sum(x => x.Col3).ToString("N0")} d/l), " +
                        $"X E5-II: {dlgV2.Where(x => x.GoodCode == "0201004").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV2.Where(x => x.GoodCode == "0201004").Sum(x => x.Col3).ToString("N0")} d/l), " +
                        $"Do 0,001S: {dlgV2.Where(x => x.GoodCode == "0601005").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV2.Where(x => x.GoodCode == "0601005").Sum(x => x.Col3).ToString("N0")} d/l), " +
                        $"Do 0,05S-II: {dlgV2.Where(x => x.GoodCode == "0601002").Sum(x => x.Col2).ToString("N0")} d/l ({dlgV2.Where(x => x.GoodCode == "0601002").Sum(x => x.Col3).ToString("N0")} d/l)";

                    var lstMarketV1 = _dbContext.TblMdMarket.Where(x => x.LocalCode == "V1").ToList();
                    foreach (var i in lstMarketV1)
                    {
                        var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.MarketCode == i.Code).Where(x => x.IsActive == true).ToList();
                        foreach (var c in lstCusPhone)
                        {
                            var info = new TblNotifySms()
                            {
                                Id = Guid.NewGuid().ToString(),
                                PhoneNumber = c.Phone,
                                Subject = Template.Title ?? "",
                                Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV1),
                                IsSend = "C",
                                Status = "TBGBL",
                                CustomerCode = c.CustomerCode,
                                MarketCode = i.Code,
                                NumberRetry = 0,
                                HeaderId = headerId
                            };
                            _dbContext.TblCmNotifySms.Add(info);
                        }
                    }
                    var lstMarketV2 = _dbContext.TblMdMarket.Where(x => x.LocalCode == "V2").ToList();
                    foreach (var i in lstMarketV2)
                    {
                        var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.MarketCode == i.Code).Where(x => x.IsActive == true).ToList();
                        foreach (var c in lstCusPhone)
                        {
                            var info = new TblNotifySms()
                            {
                                Id = Guid.NewGuid().ToString(),
                                PhoneNumber = c.Phone,
                                Subject = Template.Title ?? "",
                                Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV1),
                                IsSend = "C",
                                CustomerCode = c.CustomerCode,
                                MarketCode = i.Code,
                                Status = "TBGBL",
                                NumberRetry = 0,
                                HeaderId = headerId
                            };
                            _dbContext.TblCmNotifySms.Add(info);
                        }
                    }

                    foreach (var i in data.Pl2)
                    {
                        if (i.CustomerCode != null)
                        {
                            if (i.LocalCode == "V1")
                            {
                                var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).ToList();

                                foreach (var c in lstCusPhone)
                                {
                                    var info = new TblNotifySms()
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        PhoneNumber = c.Phone,
                                        Subject = Template.Title ?? "",
                                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV1),
                                        IsSend = "C",
                                        CustomerCode = i.CustomerCode,
                                        MarketCode = i.MarketCode,
                                        Status = "TBGBL",
                                        NumberRetry = 0,
                                        HeaderId = headerId
                                    };
                                    _dbContext.TblCmNotifySms.Add(info);
                                }
                            }
                            else
                            {
                                var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).ToList();

                                foreach (var c in lstCusPhone)
                                {
                                    var info = new TblNotifySms()
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        PhoneNumber = c.Phone,
                                        Subject = Template.Title ?? "",
                                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV2),
                                        IsSend = "C",
                                        CustomerCode = i.CustomerCode,
                                        MarketCode = i.MarketCode,
                                        Status = "TBGBL",
                                        NumberRetry = 0,
                                        HeaderId = headerId
                                    };
                                    _dbContext.TblCmNotifySms.Add(info);
                                }
                            }
                        }

                    }

                    foreach (var i in data.Pl3)
                    {
                        if (i.CustomerCode != null)
                        {
                            if (i.LocalCode == "V1")
                            {
                                var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).ToList();

                                foreach (var c in lstCusPhone)
                                {
                                    var info = new TblNotifySms()
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        PhoneNumber = c.Phone,
                                        Subject = Template.Title ?? "",
                                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV1),
                                        IsSend = "C",
                                        CustomerCode = i.CustomerCode,
                                        MarketCode = i.MarketCode,
                                        Status = "TBGBL",
                                        NumberRetry = 0,
                                        HeaderId = headerId
                                    };
                                    _dbContext.TblCmNotifySms.Add(info);
                                }
                            }
                            else
                            {
                                var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).ToList();

                                foreach (var c in lstCusPhone)
                                {
                                    var info = new TblNotifySms()
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        PhoneNumber = c.Phone,
                                        Subject = Template.Title ?? "",
                                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV2),
                                        IsSend = "C",
                                        CustomerCode = i.CustomerCode,
                                        MarketCode = i.MarketCode,
                                        Status = "TBGBL",
                                        NumberRetry = 0,
                                        HeaderId = headerId
                                    };
                                    _dbContext.TblCmNotifySms.Add(info);
                                }
                            }
                        }
                    }

                    foreach (var i in data.Pl4)
                    {
                        if (i.CustomerCode != null)
                        {
                            if (i.LocalCode == "V1")
                            {
                                var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).ToList();

                                foreach (var c in lstCusPhone)
                                {

                                    var info = new TblNotifySms()
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        PhoneNumber = c.Phone,
                                        Subject = Template.Title ?? "",
                                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV1),
                                        IsSend = "C",
                                        CustomerCode = i.CustomerCode,
                                        MarketCode = i.MarketCode,
                                        Status = "TBGBL",
                                        NumberRetry = 0,
                                        HeaderId = headerId
                                    };
                                    _dbContext.TblCmNotifySms.Add(info);
                                }
                            }
                            else
                            {
                                var lstCusPhone = _dbContext.TblMdCustomerPhone.Where(x => x.CustomerCode == i.CustomerCode).Where(x => x.IsActive == true).ToList();

                                foreach (var c in lstCusPhone)
                                {

                                    var info = new TblNotifySms()
                                    {
                                        Id = Guid.NewGuid().ToString(),
                                        PhoneNumber = c.Phone,
                                        Subject = Template.Title ?? "",
                                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay).Replace("[goods]", goodsV2),
                                        IsSend = "C",
                                        CustomerCode = i.CustomerCode,
                                        MarketCode = i.MarketCode,
                                        Status = "TBGBL",
                                        NumberRetry = 0,
                                        HeaderId = headerId
                                    };
                                    _dbContext.TblCmNotifySms.Add(info);
                                }
                            }
                        }
                    }
                }

                else
                {
                    var lstCustomerPhone = _dbContext.TblMdCustomerPhone.Where(x => x.IsActive == true).ToList();
                    foreach (var cusPhone in lstCustomerPhone)
                    {
                        var info = new TblNotifySms()
                        {
                            Id = Guid.NewGuid().ToString(),
                            PhoneNumber = cusPhone.Phone,
                            Subject = Template.Title ?? "",
                            Contents = Template.HtmlSource,
                            CustomerCode = cusPhone.CustomerCode ?? "_",
                            MarketCode = cusPhone.MarketCode ?? "_",
                            IsSend = "C",
                            Status = "TB",
                            NumberRetry = 0,
                            HeaderId = headerId
                        };
                        _dbContext.TblCmNotifySms.Add(info);
                    }
                }

                _dbContext.SaveChanges();
                return "true";
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return null;
            }
        }

        public async Task SendEmail(string headerId)
        {
            var lstCusEmail = _dbContext.TblMdCustomerEmail.Where(x => x.IsActive == true).ToList();
            var Template = _dbContext.TblAdConfigTemplate.Where(x => x.Name == "SMS thông báo giá bán lẻ").FirstOrDefault();
            try
            {
                DateTime Date = DateTime.Now;
                var Ngay = $"Từ {Date.Hour:D2}h ngày {Date:dd/MM/yyyy}";

                foreach (var item in lstCusEmail)
                {
                    var info = new TblNotifyEmail()
                    {
                        Id = Guid.NewGuid().ToString(),
                        Email = item.Email,
                        Subject = Template.Title,
                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay),
                        IsSend = "N",
                        NumberRetry = 0,
                        HeaderId = headerId
                    };
                    _dbContext.TblCmNotifiEmail.Add(info);
                }
                _dbContext.SaveChanges();

            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
            }
        }

        public async Task<List<TblNotifySms>> GetHistorySms(string headerID)
        {
            try
            {
                var data = await _dbContext.TblCmNotifySms.Where(x => x.HeaderId == headerID && x.IsSend != "K").ToListAsync();
                return data;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new List<TblNotifySms>();
            }
        }
   

        public async Task<List<NotifyEmailViewModel>> GetHistoryMail(string headerID)
        {
            try
            {
                
                var BbdoCus = _dbContext.TblMdCustomerBbdo.ToList();
                var fileDowload = _dbContext.TblBuHistoryDownload.Where(x => x.HeaderCode == headerID).ToList();
                var emailList = _dbContext.TblCmNotifiEmail
                    .Where(x => x.HeaderId == headerID && x.IsSend != "K")
                    .ToList();

                var data = emailList.Select(x => new NotifyEmailViewModel
                {
                    Id = x.Id,
                    HeaderID = x.HeaderId,
                    Status = x.Status,
                    Subject = x.Subject,
                    Contents = x.Contents,
                    NumberRetry = x.NumberRetry,
                    IsSend = x.IsSend,
                    Email = x.Email,
                    CustomerCode = x.CustomerCode,
                    CustomerName = BbdoCus.FirstOrDefault(y => y.Code == x.CustomerCode)?.Name,
                    CheckFile = fileDowload.Any(y => y.CustomerCode == x.CustomerCode) ? "có File đính kèm" : "chưa có file đính kèm"
                }).ToList();

                return data;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new List<NotifyEmailViewModel>();
            }
        }
       

        public async Task SendSMS(List<string> lstSms)
        {
            try
            {
                foreach (var i in lstSms)
                {
                    var sms = _dbContext.TblCmNotifySms.Where(x => x.Id == i).FirstOrDefault();
                    sms.IsSend = "N";
                    _dbContext.SaveChanges();
                }
            }
            catch (Exception ex)
            {
              
            }
        }

        public async Task ResetSendSMS(List<string> lstSms)
        {
            try
            {
                foreach (var i in lstSms)
                {
                    var sms = _dbContext.TblCmNotifySms.Where(x => x.Id == i).FirstOrDefault();
                    if ((sms.NumberRetry == 3 && sms.IsSend == "N") || sms.IsSend == "Y")
                    {
                        sms.IsSend = "N";
                        sms.NumberRetry = 0;

                        _dbContext.SaveChanges();
                    }
                    else
                    {

                    }

                }
            }
            catch (Exception ex)
            {

            }
        }
        public async Task SendlstMail(List<string> lstEmail)
        {
            try
            {
                foreach (var i in lstEmail)
                {
                    var sms = _dbContext.TblCmNotifiEmail.Where(x => x.Id == i).FirstOrDefault();
                    sms.IsSend = "N";
                    _dbContext.SaveChanges();
                }
            }
            catch (Exception ex)
            {

            }
        }

        public async Task ResetSendlstMail(List<string> lstEmail)
        {
            try
            {
                foreach (var i in lstEmail)
                {
                    var mail = _dbContext.TblCmNotifiEmail.Where(x => x.Id == i).FirstOrDefault();
                    if (mail.IsSend == "Y" || (mail.IsSend == "N" && mail.NumberRetry == 3))
                    {
                        mail.IsSend = "C";
                        mail.NumberRetry = 0;

                        _dbContext.SaveChanges();
                    }
                }
            }
            catch (Exception ex)
            {

            }
        }

        public async Task DeleSMS(string headerId, string status)
        {
            try
            {
                var lstSms = _dbContext.TblCmNotifySms.Where(x => x.HeaderId == headerId && x.Status == status).ToList();
                foreach (var i in lstSms)
                {
                    i.IsSend = "K";
                    _dbContext.SaveChanges();
                }
            }
            catch (Exception ex)
            {
                //return null;
            }
        }
        #endregion

        #region xử lý quy trình xét duyệt

        public async Task HandleQuyTrinh(QuyTrinhModel data)
        {
            try
            {
                data.header.Status = data.Status.Code == "06" ? "01" : data.Status.Code == "07" ? "01" : data.Status.Code;
                //data.header.Status = data.Status.Code == "09" ? "01" : data.Status.Code == "11" ? "01" : data.Status.Code == "10" ? "04" : data.Status.Code == "12" ? "04" : data.Status.Code;
                _dbContext.TblBuCalculateDiscount.Update(data.header);
                var TpkdId = _dbContext.TblAdAccountGroup.FirstOrDefault(x => x.Name == "G_TP_KD").Id;
                var AccoundTPKD = _dbContext.TblAdAccount_AccountGroup.Where(x => x.GroupId == TpkdId).ToList();
                var templateEmail = _dbContext.TblAdConfigTemplate.FirstOrDefault(x => x.Name == "Email Thông báo phê duyệt");
                var Account = _dbContext.TblAdAccount.Select(x => new { Email= x.Email, UserName = x.UserName });
                
                var h = new TblBuHistoryAction()
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = data.header.Id,

                    Action = data.Status.Code == "02" ? "Trình duyệt" : data.Status.Code == "03" ? "Yêu cầu chỉnh sửa" : data.Status.Code == "04" ? "Phê duyệt" : data.Status.Code == "05" ? "Từ chối" : data.Status.Code == "06" ? "Hủy trình duyệt" : "Hủy phê duyệt",
                    //Action = data.Status.Code == "02" ? "Trình duyệt giá bán lẻ" : data.Status.Code == "03" ? "Yêu cầu chỉnh sửa giá bán lẻ" : data.Status.Code == "04" ? "Duyệt giá bán lẻ" : data.Status.Code == "05" ? "Từ chối": data.Status.Code == "06" ? "Trình duyệt giá thù lao" : data.Status.Code == "07" ? "Yêu cầu chỉnh sửa giá thù lao" : data.Status.Code == "08" ? "Phê duyệt Đợt tính thù lao" : data.Status.Code == "09" ? "Hủy trình duyệt giá bán lẻ" : data.Status.Code == "10" ? "Hủy trình duyệt giá thù lao" : data.Status.Code == "11" ? "Hủy phê duyệt giá bán lẻ" : data.Status.Code == "12" ? "Hủy phê duyệt giá thù lao" : "Hủy phê duyệt",
                    Contents = data.Status.Content
                };
                _dbContext.TblBuHistoryAction.Add(h);
                //if (data.Status.Code == "02" || data.Status.Code == "06")
                if (data.Status.Code == "02")
                    {
                    var email = new TblNotifyEmail();
                    foreach (var i in AccoundTPKD)
                    {
                        var m = Account.FirstOrDefault(x => x.UserName == i.UserName).Email;
                        email = new TblNotifyEmail()
                        {
                            Id = Guid.NewGuid().ToString(),
                            HeaderId = data.header.Id,
                            Email =m,
                            NumberRetry = 0,
                            Subject = templateEmail.Title.Replace("[fromDate]", data.header.Date.ToString("dd/MM/yyyy")),
                            Contents = templateEmail.HtmlSource.Replace("[pram]", data.Status.Link).Replace("[fromDate]", data.header.Date.ToString("dd/MM/yyyy")),
                            IsSend = "N",
                        };
                        _dbContext.TblCmNotifiEmail.Add(email);
                    }
                }else if (data.Status.Code == "04")
                {
                    await SaveMailTBGia(data.header.Id);
                }else if (data.Status.Code == "07")
                {
                    var lstmail = _dbContext.TblCmNotifiEmail.Where(x => x.HeaderId == data.header.Id);
                    foreach (var item in lstmail)
                    {
                        item.IsSend = "K";

                    }
                    _dbContext.TblCmNotifiEmail.UpdateRange(lstmail);
                }
                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }



        #endregion

        #region
        public async Task<List<TblBuHistoryAction>> GetHistoryAction(string code)
        {
            try
            {
                var data = await _dbContext.TblBuHistoryAction.Where(x => x.HeaderCode == code).OrderByDescending(x => x.CreateDate).ToListAsync();
                return data;
            }
            catch (Exception ex)
            {
                return new List<TblBuHistoryAction>();
            }
        }
        #endregion

        #region history dowload file
        public async Task<List<TblBuHistoryDownload>> GetHistoryFile(string code)
        {
            try
            {
                var data = await _dbContext.TblBuHistoryDownload.Where(x => x.HeaderCode == code).OrderByDescending(x => x.CreateDate).ToListAsync();
                return data;
            }
            catch (Exception ex)
            {
                return new List<TblBuHistoryDownload>();
            }
        }
        //public async Task ResendEmail(string code)
        //{
        //    try
        //    {
        //        var data = _dbContext.TblCmNotifiEmail.Where(x => x.HeaderId == code&&x.IsSend=="N" && x.NumberRetry==3).ToList();

        //        foreach(var mail in data)
        //        {
        //            mail.IsSend = "N";
        //            mail.Status = "";
        //            mail.NumberRetry = 0;
                    
        //        }
        //        _dbContext.TblCmNotifiEmail.UpdateRange(data);
        //        _dbContext.SaveChanges();
               
        //    }
        //    catch (Exception ex)
        //    {
        //        this.Status = false;
        //    }
        //}
        #endregion

        #region copy kỳ
        public async Task<CalculateDiscountInputModel> CopyInput(string headerId, string id)
        {
            try 
            {
                var Header = _dbContext.TblBuCalculateDiscount.Find(headerId);
                Header.Id = id;
                Header.Date = DateTime.Now;
                Header.Status = "01";
                Header.IsActive = true;
                //var headerId = Guid.NewGuid().ToString();
                var InputPrice = await _dbContext.TblBuInputPrice.Where(x => x.HeaderId == headerId).OrderBy(x => x.Order).ToListAsync();
                //InputPrice.he
                var lstGoods = await _dbContext.TblMdGoods.OrderBy(x => x.CreateDate).OrderBy(x => x.Order).ToListAsync();
                var lstMarket = await _dbContext.TblMdMarket.OrderBy(x => x.Code).ToListAsync();
                var lstCustomerDb = await _dbContext.TblMdCustomerDb.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerPt = await _dbContext.TblMdCustomerPt.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerFob = await _dbContext.TblMdCustomerFob.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerTnpp = await _dbContext.TblMdCustomerTnpp.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerBbdo = await _dbContext.TblMdCustomerBbdo.OrderBy(x => x.Order).ToListAsync();
                var lstCustomerPts = await _dbContext.TblMdCustomerPts.OrderBy(x => x.Order).ToListAsync();

                return new CalculateDiscountInputModel
                {
                    Header = Header,
                    InputPrice = InputPrice.Select(g => new TblBuInputPrice
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = id,
                        GoodCode = g.GoodCode,
                        GoodName = g.GoodName,
                        ThueBvmt = g.ThueBvmt,
                        Vcf = g.Vcf,
                        ChenhLech = g.ChenhLech,
                        GblV1 = g.GblV1,
                        GblV2 = g.GblV2,
                        L15Blv2 = g.L15Blv2,
                        L15Nbl = g.L15Nbl,
                        LaiGop = 0,
                        IsActive = true,
                        Order = g.Order,
                    }).ToList(),
                    Market = lstMarket.Select(x => new TblBuInputMarket
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = id,
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
                        HeaderId = id,
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
                    CustomerPt = lstCustomerPt.Select(x => new TblBuInputCustomerPt
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = id,
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
                    CustomerPts = lstCustomerPts.Select(x => new TblBuInputCustomerPts
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = id,
                        Code = x.Code,
                        Name = x.Name,
                        CuLyBq = x.CuLyBq,
                        Cpccvc = x.Cpccvc,
                        Cvcbq = x.Cvcbq,
                        Lvnh = x.Lvnh,
                        Htcvc = x.Htcvc,
                        HttVb1370 = x.HttVb1370,
                        Ckv2 = x.Ckv2,
                        GoodsCode = x.GoodsCode,
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
                        HeaderId = id,
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
                    CustomerTnpp = lstCustomerTnpp.Select(x => new TblBuInputCustomerTnpp
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = id,
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
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
                    CustomerBbdo = lstCustomerBbdo.Select(x => new TblBuInputCustomerBbdo
                    {
                        Id = Guid.NewGuid().ToString(),
                        HeaderId = id,
                        Code = x.Code,
                        Name = x.Name,
                        LocalCode = x.LocalCode ?? "-",
                        DeliveryPoint = x.DeliveryPoint ?? "-",
                        DeliveryGroupCode = x.DeliveryGroupCode ?? "-",
                        GoodsCode = x.GoodsCode ?? "-",
                        MarketCode = x.MarketCode ?? "-",
                        CuLyBq = x.CuLyBq ?? 0,
                        Cpccvc = x.Cpccvc ?? 0,
                        Cvcbq = x.Cvcbq ?? 0,
                        Lvnh = x.Lvnh ?? 0,
                        Fob = x.Fob ?? 0,
                        Htcvc = x.Htcvc ?? 0,
                        HttVb1370 = x.HttVb1370 ?? 0,
                        Ckv2 = x.Ckv2 ?? 0,
                        PhuongThuc = x.PhuongThuc ?? "-",
                        Thtt = x.Thtt ?? "-",
                        Order = x.Order,
                        Adrress = x.Adrress ?? "-",
                        CkDau = x.CkDau ?? 0,
                        CkXang = x.CkXang ?? 0,
                        IsActive = true,
                        LamTronDacBiet = x.LamTronDacBiet
                    }).ToList(),
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

        #region gửi mail thông báo giá xăng dầu
        public async Task SaveMailTBGia(string headerId)
        {
            try
            {
                //var s = new ExportWordService(_dbContext);
                //var data = await this.CalculateDiscountOutput(headerId);
                var litCustomerBBdoMail = _dbContext.TblMdCustomerEmail.Where(x => x.IsActive == true && x.Email != "").ToList();
                var lstCustomerBbdo = _dbContext.TblMdCustomerBbdo.Where(x => x.IsActive == true).ToList();
                var dataHeader = await this.GetInput(headerId);
                DateTime Date = dataHeader.Header.Date;
                var Ngay = $"{Date.Hour:D2}h00 ngày {Date:dd/MM/yyyy}";
                var template = _dbContext.TblAdConfigTemplate.FirstOrDefault(x => x.Name == "Email thông báo giá bán");

                var lstmail = new List<TblNotifyEmail>();

                foreach (var item in litCustomerBBdoMail)
                {
                    if (lstCustomerBbdo.FirstOrDefault(x => x.Code == item.CustomerCode) != null)
                    {
                        lstmail.Add(new TblNotifyEmail()
                        {
                            Id = Guid.NewGuid().ToString(),
                            Email = item.Email,
                            Subject = template.Title.Replace("[fromDate]", Ngay) ?? "",
                            Contents = template.HtmlSource.Replace("[fromDate]", Ngay),
                            IsSend = "C",
                            NumberRetry = 0,
                            HeaderId = headerId,
                            CustomerCode = item.CustomerCode ?? ""
                        });
                    }
                }

                _dbContext.TblCmNotifiEmail.AddRange(lstmail);





                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }
        public async Task DelMailPheDuyet(string headerId)
        {
            try
            {
                var lstmail = _dbContext.TblCmNotifiEmail.Where(x=>x.HeaderId==headerId);
                foreach (var item in lstmail)
                {
                    item.IsSend = "K";

                }
                ;
                
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
            }
        }

        #endregion
    }
}

