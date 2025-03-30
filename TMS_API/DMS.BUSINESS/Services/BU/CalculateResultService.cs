using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.CORE.Entities.BU;
using DMS.CORE;
using System.Diagnostics;
using DMS.CORE.Entities.MD;
using DMS.BUSINESS.Dtos.MD;
using Microsoft.AspNetCore.Http;
using DMS.BUSINESS.Models;
using Microsoft.EntityFrameworkCore;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using SMO;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;
using PROJECT.Service.Extention;
using Aspose.Words;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using DocumentFormat.OpenXml;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using System.Linq;
using NPOI.HPSF;
using NPOI.SS.Formula.Functions;
using DMS.CORE.Entities.IN;
using Aspose.Words.Tables;
using System.Data;
using System.Globalization;
using NPOI.SS.Util;
using System.Net.Http.Headers;
using Microsoft.Data.SqlClient;
using System.Net.Mail;
using System.Net;
using System.Data.SqlClient;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Office.CustomUI;
using NPOI.HSSF.Record.Chart;
using System.IO.Packaging;
using OfficeOpenXml;
using FontSize = DocumentFormat.OpenXml.Wordprocessing.FontSize;
using static System.Runtime.InteropServices.JavaScript.JSType;
using TopBorder = DocumentFormat.OpenXml.Wordprocessing.TopBorder;
using BottomBorder = DocumentFormat.OpenXml.Wordprocessing.BottomBorder;
using LeftBorder = DocumentFormat.OpenXml.Wordprocessing.LeftBorder;
using RightBorder = DocumentFormat.OpenXml.Wordprocessing.RightBorder;

namespace DMS.BUSINESS.Services.BU
{
    public interface ICalculateResultService : IGenericService<TblMdGoods, GoodsDto>
    {
        Task<CalculateResultModel> GetResult(string code, int tab);
        Task<InsertModel> GetDataInput(string code);
        Task<List<TblBuHistoryAction>> GetHistoryAction(string code);
        Task<List<TblBuHistoryDownload>> GetHistoryFile(string code);
        Task<List<TblMdCustomer>> GetCustomer();
        Task UpdateDataInput(InsertModel model);
        Task SendEmail(string headerId);
        Task SendSMS(string headerId);
        Task<List<TblNotifyEmail>> GetMail(string headerId);
        Task<List<TblNotifySms>> GetSms(string headerId);
        Task<string> SaveFileHistory(MemoryStream outFileStream, string headerId);
        Task<string> GenarateWordTrinhKy(string headerId, string nameTeam);
        Task<string> GenarateWord(List<string> lstCustomerChecked, string headerId);
        Task<string> GenarateFile(List<string> lstCustomerChecked, string type, string headerId, CalculateResultModel data);
        Task<string> ExportExcelTrinhKy(string headerId);
        Task<string> ExportExcelPlus(string headerId, CalculateResultModel data);
    }
    public class CalculateResultService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdGoods, GoodsDto>(dbContext, mapper), ICalculateResultService
    {
        public async Task<CalculateResultModel> GetResult(string code, int tab)
        {
            try
            {
                var data = new CalculateResultModel();
                //var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.CreateDate).ToListAsync();
                var lstGoods = await _dbContext.TblMdGoods.Where(x => x.IsActive == true).OrderBy(x => x.CreateDate).ToListAsync();
                data.lstGoods = lstGoods;
                var lstMarket = await _dbContext.TblMdMarket.OrderBy(x => x.Code).ToListAsync();
                var lstCustomer = await _dbContext.TblMdCustomer.Where(x => x.IsActive == true).ToListAsync();
                var lstCR = await _dbContext.TblBuCalculateResultList.OrderBy(x => x.FDate).ToListAsync();

                //data.HEADER_CR = lstCR.FirstOrDefault(x => x.Code == code);

                DateTime fDate = lstCR.FirstOrDefault(x => x.Code == code).FDate;

                var OldCalculate = await _dbContext.TblBuCalculateResultList
                                                    .Where(x => x.FDate < fDate) // Lọc các trường có FDate nhỏ hơn ngày hiện tại
                                                    .Where(x => x.Status == "04") //lấy trường đã đươc phê duyệt
                                                    .OrderByDescending(x => x.FDate)
                                                    //.Select(x => x.Code) // Chọn trường Code
                                                    .FirstOrDefaultAsync();

                data.DLG.NameOld = OldCalculate.Name ?? "";
                var dataVCLOld = await _dbContext.TblInVinhCuaLo.Where(x => x.HeaderCode == OldCalculate.Code).ToListAsync();
                var dataHSMHOld = await _dbContext.TblInHeSoMatHang.Where(x => x.HeaderCode == OldCalculate.Code).ToListAsync();
                var dataVCL = await _dbContext.TblInVinhCuaLo.Where(x => x.HeaderCode == code).ToListAsync();
                var dataHSMH = await _dbContext.TblInHeSoMatHang.Where(x => x.HeaderCode == code).ToListAsync();
                var mappingBBDO = await _dbContext.TblMdMapPointCustomerGoods.Where(x => x.IsActive == true).ToListAsync();
                var dataPoint = await _dbContext.TblMdDeliveryPoint.ToListAsync();
                var lstSpecialCustomer = await _dbContext.TbLInCustomerFob.Where(x => x.HeaderCode == code).ToListAsync();
                if (dataVCL.Count() == 0 || dataHSMH.Count() == 0)
                {
                    return data;
                }
                #region DLG
                var _oDlg = 1;
                foreach (var g in lstGoods)
                {
                    var vcl = dataVCL.Where(x => x.GoodsCode == g.Code).ToList();
                    var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    data.DLG.Dlg_1.Add(new DLG_1
                    {
                        Code = g.Code,
                        Col1 = g.Name,
                        Col2 = vcl.Sum(x => x.GblcsV1),
                        Col3 = vcl.Sum(x => x.GblV2),
                        Col4 = vcl.Sum(x => x.V2_V1),
                        Col5 = vcl.Sum(x => x.MtsV1),
                        Col6 = vcl.Sum(x => x.Gny),
                        Col7 = vcl.Sum(x => x.Clgblv),
                    });

                    data.DLG.Dlg_2.Add(new DLG_2
                    {
                        Code = g.Code,
                        Col1 = g.Name,
                        Col2 = vcl.Sum(x => x.GblV2),
                    });

                    var _dlg3 = new DLG_3
                    {
                        Code = g.Code,
                        ColA = _oDlg.ToString(),
                        ColB = g.Name,
                        Col1 = hsmh.Sum(x => x.HeSoVcf),
                        Col2 = hsmh.Sum(x => x.ThueBvmt),
                        Col3 = hsmh.Sum(x => x.L15ChuaVatBvmt),
                        Col4 = hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmt),
                        Col5 = (hsmh.Sum(x => x.ThueBvmt) + hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmt)) * 1.1M,
                        Col6 = vcl.Sum(x => x.GblcsV1),
                        Col7 = vcl.Sum(x => x.GblV2),
                    };
                    if (_dlg3.Col7 != 0)
                    {
                        _dlg3.Col8 = _dlg3.Col7 / 1.1M - _dlg3.Col2;
                    }
                    _dlg3.Col9 = _dlg3.Col7 - _dlg3.Col5;
                    _dlg3.Col10 = _dlg3.Col8 - _dlg3.Col4;
                    data.DLG.Dlg_3.Add(_dlg3);

                    var _dlg5 = new DLG_5
                    {
                        Code = g.Code,
                        ColA = _oDlg.ToString(),
                        ColB = g.Name,
                        Col1 = hsmh.Sum(x => x.HeSoVcf),
                        Col2 = hsmh.Sum(x => x.ThueBvmt),
                        Col3 = hsmh.Sum(x => x.L15ChuaVatBvmt),
                        Col4 = hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmt),
                        Col5 = (hsmh.Sum(x => x.ThueBvmt) + hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmt)) * 1.1M,
                    };
                    data.DLG.Dlg_5.Add(_dlg5);

                    var _dlg6 = new DLG_6
                    {
                        Code = g.Code,
                        ColA = _oDlg.ToString(),
                        ColB = g.Name,
                        Col1 = hsmh.Sum(x => x.HeSoVcf),
                        Col2 = hsmh.Sum(x => x.ThueBvmt),
                        Col4 = hsmh.Sum(x => x.L15ChuaVatBvmt),
                        Col5 = 0,

                    };
                    if (_dlg6.Col1 != 0 && _dlg6.Col2 != 0)
                    {
                        _dlg6.Col3 = _dlg6.Col2 / _dlg6.Col1;
                    }
                    _dlg6.Col6 = _dlg6.Col4 + _dlg6.Col3;
                    _dlg6.Col7 = _dlg6.Col6 * 1.1M;
                    data.DLG.Dlg_6.Add(_dlg6);

                    _oDlg++;
                }

                data.DLG.Dlg_4.Add(new DLG_4
                {
                    ColA = "I",
                    ColB = "Vùng thị trường trung tâm",
                    IsBold = true
                });
                var _oI = 1;
                foreach (var g in lstGoods)
                {
                    var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    var dlg1 = data.DLG.Dlg_1.Where(x => x.Code == g.Code).ToList();
                    var i = new DLG_4
                    {
                        Code = g.Code,
                        Type = "TT",
                        ColA = _oI.ToString(),
                        ColB = g.Name,
                        Col1 = hsmh.Sum(x => x.HeSoVcf),
                        Col2 = hsmh.Sum(x => x.ThueBvmt),
                        Col3 = hsmh.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col4 = hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col5 = (hsmh.Sum(x => x.ThueBvmt) + hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmtNbl)) * 1.1M,
                        Col6 = dlg1.Sum(x => x.Col6),
                        Col14 = hsmh.Sum(x => x.GiamGiaFob) - 30,
                        Col10 = hsmh.Sum(x => x.LaiGopDieuTiet) == null ? 0 : hsmh.Sum(x => x.LaiGopDieuTiet),
                    };
                    if (i.Col6 != 0)
                    {
                        i.Col7 = i.Col6 / 1.1M - i.Col2;
                    }
                    i.Col8 = i.Col6 - i.Col5;
                    if (i.Col8 != 0)
                    {
                        i.Col9 = i.Col8 / 1.1M;
                    }
                    i.Col11 = i.Col1 * i.Col10 * 1.1M;
                    i.Col13 = i.Col11 + i.Col9;
                    i.Col12 = i.Col13 * 1.1M;
                    i.Col15 = (i.Col12 - i.Col14) * i.Col1;
                    i.Col16 = i.Col12 - i.Col14;

                    data.DLG.Dlg_4.Add(i);
                    _oI++;
                }

                data.DLG.Dlg_4.Add(new DLG_4
                {
                    ColA = "II",
                    ColB = "Các vùng thị trường còn lại",
                    IsBold = true
                });
                var _oII = 1;
                foreach (var g in lstGoods)
                {
                    var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    var dlg1 = data.DLG.Dlg_3.Where(x => x.Code == g.Code).ToList();
                    var i = new DLG_4
                    {
                        Code = g.Code,
                        Type = "OTHER",
                        ColA = _oII.ToString(),
                        ColB = g.Name,
                        Col1 = hsmh.Sum(x => x.HeSoVcf),
                        Col2 = hsmh.Sum(x => x.ThueBvmt),
                        Col3 = hsmh.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col4 = hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col5 = (hsmh.Sum(x => x.ThueBvmt) + hsmh.Sum(x => x.HeSoVcf) * hsmh.Sum(x => x.L15ChuaVatBvmtNbl)) * 1.1M,
                        Col6 = dlg1.Sum(x => x.Col7),
                        Col14 = hsmh.Sum(x => x.GiamGiaFob),
                        Col10 = hsmh.Sum(x => x.LaiGopDieuTiet) == null ? 0 : hsmh.Sum(x => x.LaiGopDieuTiet),
                    };
                    if (i.Col6 != 0)
                    {
                        i.Col7 = i.Col6 / 1.1M - i.Col2;
                    }
                    i.Col8 = i.Col6 - i.Col5;
                    if (i.Col8 != 0)
                    {
                        i.Col9 = i.Col8 / 1.1M;
                    }
                    i.Col11 = i.Col1 * i.Col10 * 1.1M;
                    i.Col13 = i.Col11 + i.Col9;
                    i.Col12 = i.Col13 * 1.1M;
                    i.Col15 = (i.Col12 - i.Col14) * i.Col1;
                    i.Col16 = i.Col12 - i.Col14;
                    data.DLG.Dlg_4.Add(i);
                    _oII++;
                }

                #region

                //if (OldCalculate != null)
                //{
                foreach (var g in lstGoods)
                {
                    var hsmho = dataHSMHOld.Where(x => x.GoodsCode == g.Code).ToList();
                    var vclo = dataVCLOld.Where(x => x.GoodsCode == g.Code).ToList();
                    //var dlg1 = data.DLG.Dlg_1.Where(x => x.Code == g.Code).ToList();
                    var k = new DLG_4_Old
                    {
                        Code = g.Code,
                        Type = "TT",
                        ColB = g.Name,
                        Col1 = hsmho.Sum(x => x.HeSoVcf),
                        Col2 = hsmho.Sum(x => x.ThueBvmt),
                        Col3 = hsmho.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col4 = hsmho.Sum(x => x.HeSoVcf) * hsmho.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col5 = (hsmho.Sum(x => x.ThueBvmt) + hsmho.Sum(x => x.HeSoVcf) * hsmho.Sum(x => x.L15ChuaVatBvmtNbl)) * 1.1M,
                        Col6 = vclo.Sum(x => x.Gny),
                        Col14 = hsmho.Sum(x => x.GiamGiaFob),
                        Col10 = hsmho.Sum(x => x.LaiGopDieuTiet) == null ? 0 : hsmho.Sum(x => x.LaiGopDieuTiet),
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
                    data.DLG.Dlg_4_Old.Add(k);
                }

                foreach (var g in lstGoods)
                {
                    var hsmho = dataHSMHOld.Where(x => x.GoodsCode == g.Code).ToList();
                    var vclo = dataVCLOld.Where(x => x.GoodsCode == g.Code).ToList();
                    //var dlg1 = data.DLG.Dlg_3.Where(x => x.Code == g.Code).ToList();
                    var k = new DLG_4_Old
                    {
                        Code = g.Code,
                        Type = "OTHER",
                        ColA = _oII.ToString(),
                        ColB = g.Name,
                        Col1 = hsmho.Sum(x => x.HeSoVcf),
                        Col2 = hsmho.Sum(x => x.ThueBvmt),
                        Col3 = hsmho.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col4 = hsmho.Sum(x => x.HeSoVcf) * hsmho.Sum(x => x.L15ChuaVatBvmtNbl),
                        Col5 = (hsmho.Sum(x => x.ThueBvmt) + hsmho.Sum(x => x.HeSoVcf) * hsmho.Sum(x => x.L15ChuaVatBvmtNbl)) * 1.1M,
                        Col6 = vclo.Sum(x => x.GblV2),
                        Col14 = hsmho.Sum(x => x.GiamGiaFob),
                        Col10 = hsmho.Sum(x => x.LaiGopDieuTiet) == null ? 0 : hsmho.Sum(x => x.LaiGopDieuTiet),
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
                    data.DLG.Dlg_4_Old.Add(k);

                }
                #endregion


                // thay dổi giá bán lẻ 
                foreach (var g in lstGoods)
                {
                    var vcl = dataVCLOld.Where(x => x.GoodsCode == g.Code).ToList();
                    //var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    var dlg_1 = data.DLG.Dlg_1;
                    foreach (var n in dlg_1)
                    {
                        if (g.Code == n.Code)
                        {
                            var i = new Dlg_TDGBL
                            {
                                Code = g.Code,
                                ColA = g.Name,
                                Col1 = vcl.Sum(x => x.Gny), // lấy giá niêm yết ở kì trước
                                Col2 = n.Col6,
                                TangGiam1_2 = n.Col6 - vcl.Sum(x => x.Gny),
                                Col3 = vcl.Sum(x => x.GblV2), // lấy giá niêm yết ở kì trước
                                Col4 = n.Col3,
                                TangGiam3_4 = n.Col3 - vcl.Sum(x => x.GblV2),
                            };

                            data.DLG.Dlg_TDGBL.Add(i);
                        }

                    }
                }
                // Lãi gộp
                foreach (var g in lstGoods)
                {
                    var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    var hsmho = dataHSMHOld.Where(x => x.GoodsCode == g.Code).ToList();
                    var dlg_4 = data.DLG.Dlg_4;
                    var dlg_4_Old = data.DLG.Dlg_4_Old;
                    foreach (var n in dlg_4)
                    {
                        if (g.Code == n.Code)
                        {
                            var i = new DLG_7
                            {
                                Code = g.Code,
                                ColA = g.Name,
                                Type = n.Type,
                                Col1 = dlg_4_Old.Where(x => x.Code == g.Code).Where(x => x.Type == n.Type).Sum(x => x.Col12),
                                Col2 = n.Col12,
                                TangGiam1_2 = n.Col12 - dlg_4_Old.Where(x => x.Code == g.Code).Where(x => x.Type == n.Type).Sum(x => x.Col12),
                            };

                            data.DLG.Dlg_7.Add(i);
                        }
                    }
                }

                // Đề xuất mức giảm giá
                foreach (var g in lstGoods)
                {
                    var hsmh = dataHSMHOld.Where(x => x.GoodsCode == g.Code).ToList();
                    var dlg_4 = data.DLG.Dlg_4;
                    var dlg_4_Old = data.DLG.Dlg_4_Old;

                    foreach (var n in dlg_4)
                    {
                        if (g.Code == n.Code && n.Type == "TT")
                        {
                            var i = new DLG_8
                            {
                                Code = g.Code,
                                ColA = g.Name,
                                Type = n.Type,
                                Col1 = dlg_4_Old.Where(x => x.Code == g.Code).Where(x => x.Type == "TT").Sum(x => x.Col14) + 30,
                                Col2 = n.Col14 + 30,
                                TangGiam1_2 = n.Col14 - dlg_4_Old.Where(x => x.Code == g.Code).Where(x => x.Type == "TT").Sum(x => x.Col14),
                            };
                            data.DLG.Dlg_8.Add(i);
                        }

                    }
                }

                // thay đổi giá giao phương thức bán lẻ
                foreach (var g in lstGoods)
                {
                    var hsmh = dataHSMHOld.Where(x => x.GoodsCode == g.Code).ToList();
                    //var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    var dlg_3 = data.DLG.Dlg_3;
                    foreach (var n in dlg_3)
                    {
                        if (g.Code == n.Code)
                        {
                            var i = new Dlg_TdGgptbl
                            {
                                Code = g.Code,
                                ColA = g.Name,
                                Col1 = hsmh.Sum(x => x.L15ChuaVatBvmt), // lấy giá niêm yết ở kì trước
                                Col2 = n.Col3,
                                TangGiam1_2 = n.Col3 - hsmh.Sum(x => x.L15ChuaVatBvmt),
                            };

                            data.DLG.Dlg_TdGgptbl.Add(i);
                        }

                    }
                }

                #endregion

                #region PT
                var orderPT = 1;
                foreach (var l in lstMarket.Select(x => x.LocalCode).Distinct().ToList())
                {
                    var local = _dbContext.tblMdLocal.Find(l);
                    data.PT.Add(new PT
                    {
                        ColB = local.Name,
                        IsBold = true,
                    });
                    data.PL1.Add(new PL1
                    {
                        ColB = local.Name,
                        IsBold = true,
                    });
                    foreach (var m in lstMarket.Where(x => x.LocalCode == l).ToList())
                    {
                        var i = new PT
                        {
                            Code = m.Code,
                            ColA = orderPT.ToString(),
                            ColB = m.Name,
                            Col1 = m.Gap ?? 0,
                            Col3 = m.CPChungChuaCuocVC ?? 0 + m.CuocVCBQ ?? 0,
                            Col4 = m.CPChungChuaCuocVC ?? 0,
                            Col5 = m.CuocVCBQ ?? 0,
                            Col6 = m.CkDieuTietXang ?? 0,
                            Col7 = m.CkDieuTietDau ?? 0,
                        };
                        var _2 = i.Col3;

                        var _pl1 = new PL1
                        {
                            Code = m.Code,
                            ColA = orderPT.ToString(),
                            ColB = m.Name,
                        };
                        data.PL1.Add(_pl1);
                        foreach (var _l in lstGoods)
                        {
                            //var _c = lstLGDT.Where(x => x.MarketCode == m.Code && x.GoodsCode == _l.Code);
                            var _1 = m.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Type == "TT" && x.Code == _l.Code).Sum(x => x.Col13) : data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == _l.Code).Sum(x => x.Col13);
                            //var _1 = _c == null || _c.Count() == 0 ? 0 : _c.Sum(x => x.Price);
                            i.LG.Add(Math.Round(_1));

                            var p = m.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == _l.Code && x.Type == "TT").Sum(x => x.Col14) : data.DLG.Dlg_4.Where(x => x.Code == _l.Code && x.Type == "OTHER").Sum(x => x.Col14);
                            var d = new PT_GG
                            {
                                Code = _l.Code,
                                //VAT = p - i.Col5 * m.Coefficient + i.Col6
                                VAT = p - i.Col5 * 1.1M + (_l.Type == "X" ? i.Col6 : i.Col7)
                            };
                            d.VAT = Math.Round(d.VAT == null ? 0M : d.VAT / 10) * 10;
                            d.NonVAT = d.VAT == 0 ? 0 : d.VAT / 1.1M;
                            d.NonVAT = Math.Round(d.NonVAT);
                            i.GG.Add(d);

                            _pl1.GG.Add(d.VAT);


                            var _3 = d.NonVAT;
                            i.LN.Add(_1 - _2 - _3);

                            var _h = m.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == _l.Code && x.Type == "TT").Sum(x => x.Col6) : data.DLG.Dlg_4.Where(x => x.Code == _l.Code && x.Type == "OTHER").Sum(x => x.Col6);
                            var _d = m.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == _l.Code && x.Type == "TT").Sum(x => x.Col2) : data.DLG.Dlg_4.Where(x => x.Code == _l.Code && x.Type == "OTHER").Sum(x => x.Col2);

                            var _b = new PT_BVMT
                            {
                                Code = _l.Code,
                                VAT = Math.Round(_h - d.VAT),
                            };
                            var _nVAT = _h - d.VAT != 0 ? (_h - d.VAT) / 1.1M - _d : 0;
                            _b.NonVAT = Math.Round(_nVAT);
                            i.BVMT.Add(_b);
                        }
                        data.PT.Add(i);
                        orderPT++;
                    }
                }
                #endregion

                #region ĐB

                var _oDb = 1;
                foreach (var v in lstCustomer.Where(x => x.CustomerTypeCode == "KBM").OrderBy(x => x.LocalCode).Select(x => x.LocalCode).Distinct().ToList())
                {
                    var _v = _dbContext.tblMdLocal.Find(v);
                    data.DB.Add(new DB
                    {
                        ColB = _v.Name,
                        IsBold = true,
                    });
                    foreach (var c in lstCustomer.Where(x => x.LocalCode == v && x.CustomerTypeCode == "KBM").ToList())
                    {

                        var _c = new DB
                        {
                            Code = c.Code,
                            ColA = _oDb.ToString(),
                            ColB = c.Name,
                            Col1 = lstMarket.FirstOrDefault(x => x.Code == c.MarketCode).Name,
                            Col2 = c.Gap ?? 0,
                            Col4 = data.PT.Where(x => x.Code == c.MarketCode).Sum(x => x.Col4),
                            Col5 = c.CuocVcBq ?? 0,
                            Col6 = 0,
                            Col3 = data.PT.Where(x => x.Code == c.MarketCode).Sum(x => x.Col4 + x.Col6) + c.CuocVcBq ?? 0,
                            Col8 = 0,
                            Col9 = c.MgglhXang ??0,
                            Col10 = c.MgglhDau ?? 0,

                        };
                        var _rPt = data.PT.FirstOrDefault(x => x.Code == c.MarketCode);
                        foreach (var g in lstGoods)
                        {
                            var _lg = c.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col13) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col13);
                            _c.LG.Add(Math.Round(_lg));
                            //var vat = g.Type == "X" ? _rPt.GG.Where(x => x.Code == g.Code).Sum(x => x.VAT) + _c.Col9 + _c.Col8 : _rPt.GG.Where(x => x.Code == g.Code).Sum(x => x.VAT) + _c.Col10 + _c.Col8;
                            var vat = g.Type == "X"
                                            ? _rPt.GG.Where(x => x.Code == g.Code).Sum(x => x.VAT) + (g.Code == "601005" ? 0 : _c.Col9 + _c.Col8)
                                            : _rPt.GG.Where(x => x.Code == g.Code).Sum(x => x.VAT) + (g.Code == "601005" ? 0 : _c.Col10 + _c.Col8);

                            var nonVat = vat == 0 ? 0 : vat / 1.1M;
                            _c.GG.Add(new DB_GG
                            {
                                VAT = Math.Round(vat),
                                NonVAT = Math.Round(nonVat),
                            });

                            _c.LN.Add(Math.Round(_lg - _c.Col3 - nonVat - 0 / 1.1M));


                            var bv = c.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col6) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col6);
                            var t = c.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col2) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col2);

                            var bv_vat = bv - vat;
                            var bv_nonVat = bv_vat != 0 ? bv_vat / 1.1M - t : 0;
                            _c.BVMT.Add(new DB_BVMT
                            {
                                Code = g.Code,
                                VAT = Math.Round(bv_vat),
                                NonVAT = Math.Round(bv_nonVat)
                            });


                        }
                        ;
                        data.DB.Add(_c);
                        _oDb++;
                    }
                }
                #endregion

                #region FOB

                foreach (var l in lstCustomer.Where(x => x.CustomerTypeCode == "KBB").OrderBy(x => x.LocalCode).Select(x => x.LocalCode).Distinct().ToList())
                {
                    var local = _dbContext.tblMdLocal.Find(l);
                    data.FOB.Add(new FOB
                    {
                        ColB = local.Name,
                        IsBold = true,
                    });
                    var _oFob = 1;
                    foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "KBB" && x.LocalCode == l))
                    {
                        var _fob = new FOB
                        {
                            Code = c.Code,
                            ColA = _oFob.ToString(),
                            ColB = c.Name,
                            Col2 = data.PT.Where(x => x.Code == c.MarketCode).Sum(x => x.Col4),
                            Col3 = c.CuocVcBq ?? 0,
                            Col4 = 0,
                            Col1 = c.CuocVcBq ?? 0 + data.PT.Where(x => x.Code == c.MarketCode).Sum(x => x.Col4),
                            Col5 = 0,
                            Col6 = c.MgglhXang ?? 0,
                            Col7 = c.MgglhDau ?? 0,
                            Col8 = 0,
                        };
                        _oFob++;
                        var lPT = data.PT.FirstOrDefault(x => x.Code == c.MarketCode);
                        foreach (var g in lstGoods)
                        {
                            var lg = l == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col13) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col13);
                            _fob.LG.Add(Math.Round(lg));

                            var gg = l == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col14) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col14);
                            var vat = g.Type == "X" ? gg + _fob.Col6 + _fob.Col5 : gg + _fob.Col7 + _fob.Col5;
                            var nonVat = vat == 0 ? 0 : vat / 1.1M;

                            _fob.GG.Add(new FOB_GG
                            {
                                VAT = Math.Round(vat),
                                NonVAT = Math.Round(nonVat),
                            });
                            _fob.LN.Add(Math.Round(lg - _fob.Col1 - nonVat - _fob.Col8 / 1.1M));


                            var bv = c.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col6) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col6);
                            var t = c.LocalCode == "V1" ? data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col2) : data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col2);

                            var bv_vat = bv - vat;
                            var bv_nonVat = bv_vat != 0 ? bv_vat / 1.1M - t : 0;
                            _fob.BVMT.Add(new PT_BVMT
                            {
                                Code = g.Code,
                                VAT = Math.Round(bv_vat),
                                NonVAT = Math.Round(bv_nonVat)
                            });

                        }
                        data.FOB.Add(_fob);

                    }
                }

                #endregion

                #region PT09
                var _oPt09 = 1;
                foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "TNPP").ToList())
                {
                    var _i = new PT09
                    {
                        Code = c.Code,
                        ColA = _oPt09.ToString(),
                        ColB = c.Name,
                        Col5 = 0,
                        Col6 = 0,
                        Col4 = lstMarket.FirstOrDefault()?.CPChungChuaCuocVC ?? 0,
                        Col7 = c.MgglhXang ?? 0,
                        Col8 = c.MgglhDau ?? 0,
                        Col18 = 0,
                    };
                    _i.Col3 = _i.Col4 + _i.Col5 + _i.Col6;

                    var _pl4 = new PL4
                    {
                        Code = c.Code,
                        ColA = _oPt09.ToString(),
                        ColB = c.Name,
                    };
                    data.PL4.Add(_pl4);


                    foreach (var g in lstGoods)
                    {
                        var _2 = Math.Round(data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col13));
                        _i.LG.Add(_2);

                        var _x_d = g.Type == "X" ? _i.Col7 : _i.Col8;
                        var vat = Math.Round(data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col14)) + _x_d;
                        var nonVat = vat != 0 ? vat / 1.1M : 0;
                        _i.GG.Add(new PT09_GG
                        {
                            VAT = Math.Round(vat),
                            NonVAT = Math.Round(nonVat),
                        });
                        _i.LN.Add(Math.Round(_2 - _i.Col3 - nonVat - _i.Col18));
                        _pl4.GG.Add(Math.Round(vat));


                        var bv = data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col6);
                        var t = data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "OTHER").Sum(x => x.Col2);

                        var bv_vat = bv - vat;
                        var bv_nonVat = bv_vat != 0 ? bv_vat / 1.1M - t : 0;
                        _i.BVMT.Add(new PT_BVMT
                        {
                            Code = g.Code,
                            VAT = Math.Round(bv_vat),
                            NonVAT = Math.Round(bv_nonVat)
                        });


                    }
                    _oPt09++;
                    data.PT09.Add(_i);

                }
                #endregion

                #region PL2
                var _oPl2 = 1;
                foreach (var l in lstMarket)
                {
                    data.PL2.Add(new PL2
                    {
                        ColB = l.Name,
                        IsBold = true,
                    });
                    foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "KBM" && x.MarketCode == l.Code).ToList())
                    {
                        var pl2 = new PL2
                        {
                            Code = c.Code,
                            ColA = _oPl2.ToString(),
                            ColB = c.Name,
                        };
                        var _db = data.DB.FirstOrDefault(x => x.Code == c.Code);
                        _oPl2++;
                        if (_db == null) continue;
                        foreach (var gg in _db.GG)
                        {
                            pl2.GG.Add(gg.VAT);
                        }
                        data.PL2.Add(pl2);
                    }
                }


                #endregion

                #region PL3
                var _oPl3 = 1;
                foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "KBB").ToList())
                {
                    var pl3 = new PL3
                    {
                        Code = c.Code,
                        ColA = _oPl3.ToString(),
                        ColB = c.Name,
                    };
                    var dataFob = data.FOB.FirstOrDefault(x => x.Code == c.Code);
                    _oPl3++;
                    if (dataFob == null) continue;
                    foreach (var gg in dataFob.GG)
                    {
                        pl3.GG.Add(gg.VAT);
                    }

                    data.PL3.Add(pl3);
                }

                #endregion

                #region VK11 PT
                foreach (var g in lstGoods)
                {
                    data.VK11PT.Add(new VK11PT
                    {
                        ColB = g.Name,
                        IsBold = true,
                    });
                    var _o = 1;
                    foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "VK11PT").ToList())
                    {
                        var m = _dbContext.TblMdMarket.Find(c.MarketCode);
                        var v = data.PT.FirstOrDefault(x => x.Code == c.MarketCode).BVMT.Where(x => x.Code == g.Code).Sum(x => x.NonVAT);
                        var _i = new VK11PT
                        {
                            ColA = _o.ToString(),
                            ColB = c.Name,
                            Col1 = m?.Name,
                            Col2 = c.Gap ?? 0,
                            Col4 = "10",
                            Col5 = c.Code,
                            Col6 = g.Code,
                            Col7 = "L",
                            Col9 = Math.Round(v),
                            Col10 = "VND",
                            Col11 = 1,
                            Col12 = "L",
                            Col13 = "C",
                            Col14 = fDate.ToString("dd.MM.yyyy"),
                            Col15 = fDate.ToString("HH:mm"),
                            Col16 = $"31.12.9999",
                        };
                        data.VK11PT.Add(_i);
                        _o++;
                    }
                }
                #endregion

                #region VK11 ĐB
                foreach (var g in lstGoods)
                {
                    data.VK11DB.Add(new VK11DB
                    {
                        ColB = g.Name,
                        IsBold = true,
                    });
                    var _o = 1;
                    foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "KBM").ToList())
                    {
                        var m = _dbContext.TblMdMarket.Find(c.MarketCode);
                        var v = data.DB.FirstOrDefault(x => x.Code == c.Code).BVMT.Where(x => x.Code == g.Code).Sum(x => x.NonVAT);
                        var _i = new VK11DB
                        {
                            ColA = _o.ToString(),
                            ColB = c.Address,
                            ColC = c.Name,
                            Col1 = m?.Name,
                            Col2 = c.Gap ?? 0,
                            Col4 = "10",
                            Col5 = c.Code,
                            Col6 = g.Code,
                            Col7 = "L",
                            Col9 = Math.Round(v),
                            Col10 = "VND",
                            Col11 = 1,
                            Col12 = "L",
                            Col13 = "C",
                            Col14 = fDate.ToString("dd.MM.yyyy"),
                            Col15 = fDate.ToString("HH:mm"),
                            Col16 = $"31.12.9999",
                        };
                        data.VK11DB.Add(_i);
                        _o++;
                    }
                }
                #endregion

                #region VK11 FOB
                foreach (var g in lstGoods)
                {
                    data.VK11FOB.Add(new VK11FOB
                    {
                        ColB = g.Name,
                        IsBold = true,
                    });
                    var _o = 1;
                    foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "KBB").ToList())
                    {
                        var m = _dbContext.TblMdMarket.Find(c.MarketCode);
                        var v = data.FOB.FirstOrDefault(x => x.Code == c.Code).BVMT.Where(x => x.Code == g.Code).Sum(x => x.NonVAT);
                        var _i = new VK11FOB
                        {
                            ColA = _o.ToString(),
                            ColB = c.Address,
                            ColC = c.Name,
                            Col1 = m?.Name,
                            Col2 = c.Gap ?? 0,
                            Col4 = "10",
                            Col5 = c.Code,
                            Col6 = g.Code,
                            Col7 = "L",
                            Col9 = Math.Round(v),
                            Col10 = "VND",
                            Col11 = 1,
                            Col12 = "L",
                            Col13 = "C",
                            Col14 = fDate.ToString("dd.MM.yyyy"),
                            Col15 = fDate.ToString("HH:mm"),
                            Col16 = $"31.12.9999",
                        };
                        data.VK11FOB.Add(_i);
                        _o++;
                    }
                }
                #endregion

                #region VK11 TNPP
                foreach (var g in lstGoods)
                {
                    data.VK11TNPP.Add(new VK11TNPP
                    {
                        ColB = g.Name,
                        IsBold = true,
                    });
                    var _o = 1;
                    foreach (var c in lstCustomer.Where(x => x.CustomerTypeCode == "TNPP").ToList())
                    {
                        var m = _dbContext.TblMdMarket.Find(c.MarketCode);
                        var v = data.PT09.FirstOrDefault(x => x.Code == c.Code).BVMT.Where(x => x.Code == g.Code).Sum(x => x.NonVAT);
                        var _i = new VK11TNPP
                        {
                            ColA = _o.ToString(),
                            ColB = c.Address,
                            ColC = c.Name,
                            Col1 = m?.Name,
                            Col2 = c.Gap ?? 0,
                            Col4 = "9",
                            Col5 = c.Code,
                            Col6 = g.Code,
                            Col7 = "L",
                            Col9 = Math.Round(v),
                            Col10 = "VND",
                            Col11 = 1,
                            Col12 = "L",
                            Col13 = "C",
                            Col14 = fDate.ToString("dd.MM.yyyy"),
                            Col15 = fDate.ToString("HH:mm"),
                            Col16 = $"31.12.9999",
                        };
                        data.VK11TNPP.Add(_i);
                        _o++;
                    }
                }
                #endregion

                #region PST
                var customerPTS = lstCustomer.SingleOrDefault(c => c.Code == "310463");
                var ptsIndex = 1;
                foreach (var g in lstGoods)
                {
                    var hsmh = dataHSMH.Where(x => x.GoodsCode == g.Code).ToList();
                    data.PTS.Add(new PTS
                    {
                        ColA = ptsIndex.ToString(),
                        Col1 = "10",
                        Col2 = customerPTS.Code,
                        Col3 = g.Code,
                        Col4 = "L15",
                        Col5 = customerPTS.PaymentTerm,
                        Col6 = hsmh.Sum(x => x.L15ChuaVatBvmt),
                        Col7 = 1,
                        Col8 = "L15",
                        Col9 = fDate.ToString("dd.MM.yyyy"),
                        Col10 = "31.12.9999",
                    });
                    ptsIndex++;
                }
                #endregion

                #region BBDO
                foreach (var g in lstGoods.OrderByDescending(x => x.CreateDate).ToList())
                {
                    var c = mappingBBDO.Where(x => x.GoodsCode == g.Code && x.Type == "BBDO").ToList();
                    if (c.Count() == 0) continue;
                    data.BBDO.Add(new BBDO
                    {
                        ColB = g.Name,
                        IsBold = true
                    });
                    var _m = new List<BBDO_MAP>();
                    foreach (var i in c)
                    {
                        _m.Add(new BBDO_MAP
                        {
                            CustomerCode = i.CustomerCode,
                            PointCode = i.DeliveryPointCode,
                            CustomerName = lstCustomer.FirstOrDefault(x => x.Code == i.CustomerCode)?.Name,
                            PointName = dataPoint.FirstOrDefault(x => x.Code == i.DeliveryPointCode)?.Name,
                            CuocVcBq = i.CuocVcBq ?? 0
                        });
                    }
                    _m = _m.OrderBy(x => x.CustomerName).ThenBy(x => x.PointName).ToList();
                    var _o = 1;
                    //var specialCustomerDict = lstSpecialCustomer.ToDictionary(x => x.CustomerCode, x => x.Fob);
                    foreach (var e in _m)
                    {
                        //bool isSpecial = specialCustomerDict.ContainsKey(e.CustomerCode);
                        //decimal col12Value = isSpecial
                        //    ? specialCustomerDict[e.CustomerCode] ?? 0
                        //    : Math.Round(data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col14) ?? 0);
                        var i = new BBDO
                        {
                            ColA = _o.ToString(),
                            ColB = e.CustomerName,
                            ColC = e.PointName,
                            ColD = g.Name,
                            Col1 = "07",
                            Col2 = e.CustomerCode,
                            Col3 = g.Code,
                            Col4 = "L",
                            Col5 = lstCustomer.FirstOrDefault(x => x.Code == e.CustomerCode)?.PaymentTerm,
                            Col6 = Math.Round(data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col12)),
                            Col9 = data.PT.FirstOrDefault(x => x.IsBold == false)?.Col4 ?? 0,
                            Col10 = e.CuocVcBq,
                            Col11 = lstCustomer.FirstOrDefault(x => x.Code == e.CustomerCode)?.BankLoanInterest ?? 0,

                            Col12 = 1,
                        };
                        i.Col8 = Math.Round((i.Col9) + (i.Col10) + (i.Col11));
                        i.Col7 = i.Col6 == 0 ? 0 : Math.Round(i.Col6 / 1.1M);
                        i.Col13 = i.Col12 == 0 ? 0 : Math.Round(i.Col12 / 1.1M );
                        i.Col14 = Math.Round((i.Col12) - (i.Col10 ) * 1.1M);
                        i.Col15 = i.Col14 == 0 ? 0 : Math.Round(i.Col14 / 1.1M );
                        i.Col17 = Math.Round(i.Col7 - i.Col8 - i.Col15 - i.Col11 );
                        i.Col16 = Math.Round(i.Col17 * 1.1M);
                        //i.Col19 = Math.Round(data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col6) - i.Col14  ?? 0);

                        //i.Col19 =  (e.CustomerCode== "305152" ||e.CustomerCode== "308417") ? ((data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col6) - i.Col16)==0 ? 0 : (int)((Math.Floor((data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col6) - i.Col16 ?? 0)- (decimal)0.5) -5)/100)*100) :(data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col6) - i.Col16 ??0) ;
                        i.Col19 = (data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col6) - i.Col14);
                        i.Col18 = i.Col19 == 0 ? 0 : Math.Round(i.Col19 / 1.1M - data.DLG.Dlg_4.Where(x => x.Type == "OTHER" && x.Code == g.Code).Sum(x => x.Col2));

                        data.BBDO.Add(i);
                        _o++;
                    }

                }
                #endregion

                #region DO FO
                var lstMapDOFO = mappingBBDO.Where(x => x.Type == "BBFO").ToList();
                var _oDOFO = 1;
                foreach (var g in lstGoods)
                {
                    foreach (var _c in lstMapDOFO.Where(x => x.GoodsCode == g.Code).Select(x => x.CustomerCode).ToList().Distinct().ToList())
                    {
                        //var g = lstGoods.FirstOrDefault(x => x.Code == lstMapDOFO.FirstOrDefault().GoodsCode);
                        var lstPointCode = lstMapDOFO.Where(x => x.CustomerCode == _c).Select(x => x.DeliveryPointCode).ToList();
                        var lstPoint = dataPoint.Where(x => lstPointCode.Contains(x.Code)).ToList();

                        data.BBFO.Add(new BBFO
                        {
                            ColA = _oDOFO.ToString(),
                            ColB = lstCustomer.FirstOrDefault(x => x.Code == _c)?.Name,
                            IsBold = true,
                        });
                        foreach (var p in lstPoint)
                        {
                            var _5 = lstMapDOFO.FirstOrDefault(x => x.CustomerCode == _c && x.DeliveryPointCode == p.Code);
                            var i = new BBFO
                            {
                                ColA = "-",
                                ColB = p.Name,
                                ColC = g.Name ?? "",
                                Col1 = Math.Round(data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col8)),
                                Col2 = Math.Round(data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col9) ),
                                Col4 = data.PT.FirstOrDefault(x => x.IsBold == false)?.Col4 ?? 0,
                                Col5 = _5?.CuocVcBq ?? 0,
                                Col6 = Math.Round(lstCustomer.FirstOrDefault(x => x.Code == _c)?.BankLoanInterest ?? 0),
                            };
                            i.Col3 = Math.Round(i.Col4 + i.Col5 + i.Col6);
                            var _8 = data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col7);
                            i.Col8 = _8 == 0 ? 0 : Math.Round(_8 / 100 );
                            i.Col7 = Math.Round(i.Col8 + i.Col3 - i.Col2);
                            i.Col10 = Math.Round(data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col6) + i.Col7);
                            i.Col9 = i.Col10 == 0 ? 0 : Math.Round(i.Col10 / 1.1M - data.DLG.Dlg_4.Where(x => x.Code == g.Code && x.Type == "TT").Sum(x => x.Col2));


                            data.BBFO.Add(i);
                        }

                        _oDOFO++;
                    }
                }

                #endregion

                #region VK11 BB
                foreach (var i in data.BBDO)
                {
                    data.VK11BB.Add(new VK11BB
                    {
                        ColA = i.ColA,
                        ColB = i.ColB,
                        ColC = i.ColC,
                        ColD = i.ColD,
                        Col1 = i.Col1,
                        Col2 = i.Col2,
                        Col3 = i.Col3,
                        Col4 = i.Col4,
                        Col5 = i.Col5,
                        Col6 = i.Col18,
                        Col7 = "VND",
                        Col8 = "1",
                        Col9 = "L",
                        Col10 = "C",
                        Col11 = fDate.ToString("dd.MM.yyyy"),
                        Col12 = fDate.ToString("HH:mm"),
                        Col13 = $"31.12.9999",
                        IsBold = i.IsBold,
                    });
                }
                #endregion

                #region Tổng hợp
                data.Summary.Add(new Summary
                {
                    Code = Guid.NewGuid().ToString(),
                    ColB = "TNQTM",
                    IsBold = true,
                });
                foreach (var i in data.VK11PT)
                {
                    data.Summary.Add(new Summary
                    {
                        Code = Guid.NewGuid().ToString(),
                        ColA = i.ColA,
                        ColB = i.ColB,
                        ColC = i.Col1,
                        Col1 = i.Col4,
                        Col2 = i.Col5,
                        Col3 = i.Col6,
                        Col4 = i.Col7,
                        Col5 = i.Col8,
                        Col6 = i.Col9,
                        Col7 = "VND",
                        Col8 = "1",
                        Col9 = "L",
                        Col10 = "C",
                        Col11 = fDate.ToString("dd.MM.yyyy"),
                        Col12 = fDate.ToString("HH:mm"),
                        Col13 = $"31.12.9999",
                        IsBold = i.IsBold,
                    });
                }

                data.Summary.Add(new Summary
                {
                    Code = Guid.NewGuid().ToString(),
                    ColB = "KHÁCH ĐẶC BIỆT",
                    IsBold = true,
                });
                foreach (var i in data.VK11DB)
                {
                    data.Summary.Add(new Summary
                    {
                        Code = Guid.NewGuid().ToString(),
                        ColA = i.ColA,
                        ColB = i.ColB,
                        ColC = i.ColC,
                        ColD = i.Col1,
                        Col1 = i.Col4,
                        Col2 = i.Col5,
                        Col3 = i.Col6,
                        Col4 = i.Col7,
                        Col5 = i.Col8,
                        Col6 = i.Col9,
                        Col7 = "VND",
                        Col8 = "1",
                        Col9 = "L",
                        Col10 = "C",
                        Col11 = fDate.ToString("dd.MM.yyyy"),
                        Col12 = fDate.ToString("HH:mm"),
                        Col13 = $"31.12.9999",
                        IsBold = i.IsBold,
                    });
                }

                data.Summary.Add(new Summary
                {
                    Code = Guid.NewGuid().ToString(),
                    ColB = "BÁN FOB",
                    IsBold = true,
                });
                foreach (var i in data.VK11FOB)
                {
                    data.Summary.Add(new Summary
                    {
                        Code = Guid.NewGuid().ToString(),
                        ColA = i.ColA,
                        ColB = i.IsBold ? i.ColB : i.ColC,
                        ColC = i.Col1,
                        Col1 = i.Col4,
                        Col2 = i.Col5,
                        Col3 = i.Col6,
                        Col4 = i.Col7,
                        Col5 = i.Col8,
                        Col6 = i.Col9,
                        Col7 = "VND",
                        Col8 = "1",
                        Col9 = "L",
                        Col10 = "C",
                        Col11 = fDate.ToString("dd.MM.yyyy"),
                        Col12 = fDate.ToString("HH:mm"),
                        Col13 = $"31.12.9999",
                        IsBold = i.IsBold,
                    });
                }

                data.Summary.Add(new Summary
                {
                    Code = Guid.NewGuid().ToString(),
                    ColB = "TNPP",
                    IsBold = true,
                });
                foreach (var i in data.VK11TNPP)
                {
                    data.Summary.Add(new Summary
                    {
                        Code = Guid.NewGuid().ToString(),
                        ColA = i.ColA,
                        ColB = i.IsBold ? i.ColB : i.ColC,
                        ColC = i.Col1,
                        Col1 = i.Col4,
                        Col2 = i.Col5,
                        Col3 = i.Col6,
                        Col4 = i.Col7,
                        Col5 = i.Col8,
                        Col6 = i.Col9,
                        Col7 = "VND",
                        Col8 = "1",
                        Col9 = "L",
                        Col10 = "C",
                        Col11 = fDate.ToString("dd.MM.yyyy"),
                        Col12 = fDate.ToString("HH:mm"),
                        Col13 = $"31.12.9999",
                        IsBold = i.IsBold,
                    });
                }

                data.Summary.Add(new Summary
                {
                    Code = Guid.NewGuid().ToString(),
                    ColB = "BÁN BUÔN",
                    IsBold = true,
                });
                foreach (var i in data.VK11BB)
                {
                    data.Summary.Add(new Summary
                    {
                        Code = Guid.NewGuid().ToString(),
                        ColA = i.ColA,
                        ColB = i.ColB,
                        ColC = i.ColC,
                        ColD = i.ColD,
                        Col1 = i.Col1,
                        Col2 = i.Col2,
                        Col3 = i.Col3,
                        Col4 = i.Col4,
                        Col5 = i.Col5,
                        Col6 = i.Col6,
                        Col7 = "VND",
                        Col8 = "1",
                        Col9 = "L",
                        Col10 = "C",
                        Col11 = fDate.ToString("dd.MM.yyyy"),
                        Col12 = fDate.ToString("HH:mm"),
                        Col13 = $"31.12.9999",
                        IsBold = i.IsBold,
                    });
                }
                #endregion

                var rData = await RoundNumberData(data);
                return rData;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new CalculateResultModel();
            }
        }

        public async Task<CalculateResultModel> ReturnTabData(CalculateResultModel rData, int tab)
        {
            var tabResult = new CalculateResultModel();
            tabResult.lstGoods = rData.lstGoods;
            switch (tab)
            {
                case 0:
                    tabResult.DLG = rData.DLG;
                    break;
                case 1:
                    tabResult.PT = rData.PT;
                    break;
                case 2:
                    tabResult.DB = rData.DB;
                    break;
                case 3:
                    tabResult.FOB = rData.FOB;
                    break;
                case 4:
                    tabResult.PT09 = rData.PT09;
                    break;
                case 5:
                    tabResult.BBDO = rData.BBDO;
                    break;
                case 6:
                    tabResult.BBFO = rData.BBFO;
                    break;
                case 7:
                    tabResult.PL1 = rData.PL1;
                    break;
                case 8:
                    tabResult.PL2 = rData.PL2;
                    break;
                case 9:
                    tabResult.PL3 = rData.PL3;
                    break;
                case 10:
                    tabResult.PL4 = rData.PL4;
                    break;
                case 11:
                    tabResult.VK11PT = rData.VK11PT;
                    break;
                case 12:
                    tabResult.VK11DB = rData.VK11DB;
                    break;
                case 13:
                    tabResult.VK11FOB = rData.VK11FOB;
                    break;
                case 14:
                    tabResult.VK11TNPP = rData.VK11TNPP;
                    break;
                case 15:
                    tabResult.PTS = rData.PTS;
                    break;
                case 16:
                    tabResult.VK11BB = rData.VK11BB;
                    break;
                case 17:
                    tabResult.Summary = rData.Summary;
                    break;
            }
            return tabResult;
        }

        public async Task<CalculateResultModel> RoundNumberData(CalculateResultModel data)
        {
            try
            {
                foreach (var i in data.DLG.Dlg_3)
                {
                    i.Col4 = Math.Round(i.Col4);
                    i.Col5 = Math.Round(i.Col5);
                    i.Col8 = Math.Round(i.Col8);
                    i.Col9 = Math.Round(i.Col9);
                    i.Col10 = Math.Round(i.Col10);
                }
                foreach (var i in data.DLG.Dlg_5)
                {
                    i.Col4 = Math.Round(i.Col4);
                    i.Col5 = Math.Round(i.Col5);
                }
                foreach (var i in data.DLG.Dlg_6)
                {
                    i.Col3 = Math.Round(i.Col3);
                    i.Col6 = Math.Round(i.Col6);
                    i.Col7 = Math.Round(i.Col7);
                }
                foreach (var i in data.DLG.Dlg_4)
                {
                    i.Col4 = Math.Round(i.Col4);
                    i.Col5 = Math.Round(i.Col5);
                    i.Col8 = Math.Round(i.Col8);
                    i.Col9 = Math.Round(i.Col9);
                    i.Col10 = Math.Round(i.Col10);
                    i.Col7 = Math.Round(i.Col7);
                    i.Col12 = Math.Round(i.Col12);
                    i.Col13 = Math.Round(i.Col13);
                    i.Col15 = Math.Round(i.Col15);
                    i.Col16 = Math.Round(i.Col16);
                }
                return data;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new CalculateResultModel();
            }
        }

        public async Task SendEmail(string headerId)

        {
            var customer = _dbContext.TblMdCustomerContact.Where(x => x.Type == "email" && x.IsActive == true);
            var Template = _dbContext.TblAdConfigTemplate.Where(x => x.Name == "Mẫu email gửi đi").FirstOrDefault();
            try
            {
                DateTime Date = DateTime.Now;
                var Ngay = $"Từ {Date.Hour}h ngày {Date:dd/MM/yyyy}";

              
                foreach (var item in customer)
                {
                    var info = new TblNotifyEmail()
                    {
                        Id = Guid.NewGuid().ToString(),
                        Email = item.Value,
                        Subject = Template.Title,
                        Contents = Template.HtmlSource.Replace("[fromDate]",Ngay),
                        IsSend = "N",
                        NumberRetry = 0,
                        HeaderId=headerId
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

        public async Task SendSMS(string headerId)

        {
            var customer = _dbContext.TblMdCustomerContact.Where(x => x.Type == "email" && x.IsActive == true);
            var Template = _dbContext.TblAdConfigTemplate.Where(x => x.Name == "SMS thông báo thu phí").FirstOrDefault();
            try
            {
                DateTime Date = DateTime.Now;
                var Ngay = $"Từ {Date.Hour}h ngày {Date:dd/MM/yyyy}";


                foreach (var item in customer)
                {
                    var info = new TblNotifySms()
                    {
                        Id = Guid.NewGuid().ToString(),
                        PhoneNumber = item.Value,
                        Subject = Template.Title,
                        Contents = Template.HtmlSource.Replace("[fromDate]", Ngay),
                        IsSend = "N",
                        NumberRetry = 0,
                        HeaderId = headerId
                    };
                    _dbContext.TblCmNotifySms.Add(info);
                }
                _dbContext.SaveChanges();

            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;

            }
        }

        public async Task<List<TblNotifySms>> GetSms(string headerID)
        {
            try {
                var data = await _dbContext.TblCmNotifySms.Where(x => x.HeaderId == headerID).ToListAsync();
                return data;
            }
            catch(Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new List<TblNotifySms>();
            }
        }

        public async Task<List<TblNotifyEmail>> GetMail(string headerID)
        {
            try
            {
                var data = await _dbContext.TblCmNotifiEmail.Where(x => x.HeaderId == headerID).ToListAsync();
                return data;
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
                return new List<TblNotifyEmail>();
            }
        }

        public async Task<InsertModel> GetDataInput(string code)
        {

            try
            {
                var data = new InsertModel();

                data.FOB = await _dbContext.TbLInCustomerFob.Where(x => x.HeaderCode == code).ToListAsync();
                if (data.FOB.Count == 0)
                {
                    var lstCustomer = await _dbContext.TblMdCustomer.Where(x => x.IsActive == true && x.Fob > 0).ToListAsync();
                    if (lstCustomer.Count > 0)
                    {
                        var obj = new InsertModel();
                        foreach (var c in lstCustomer)
                        {
                            obj.FOB.Add(new TbLInCustomerFob
                            {
                                Code = Guid.NewGuid().ToString(),
                                HeaderCode = code,
                                CustomerName = c.Name,
                                CustomerCode = c.Code,
                                Fob = c.Fob,
                                IsActive = true,
                                CreateDate = DateTime.Now,
                            });
                        }
                        _dbContext.TbLInCustomerFob.AddRange(obj.FOB);
                        _dbContext.SaveChangesAsync();
                        data.FOB = obj.FOB;
                    }
                }

                data.Header = await _dbContext.TblBuCalculateResultList.FindAsync(code);
                data.HS1 = await _dbContext.TblInHeSoMatHang.Where(x => x.HeaderCode == code).ToListAsync();
                data.HS2 = await _dbContext.TblInVinhCuaLo.Where(x => x.HeaderCode == code).ToListAsync();
                data.Status.Code = "01";
                return data;
            }
            catch (Exception ex)
            {

                return new InsertModel();
            }
        }

        public async Task UpdateDataInput(InsertModel model)
        {
            try
            {
                _dbContext.TblInHeSoMatHang.UpdateRange(model.HS1);
                _dbContext.TblInVinhCuaLo.UpdateRange(model.HS2);

                if (model.Header.Status == model.Status.Code)
                {
                    model.Header.Status = "01";
                    _dbContext.TblBuCalculateResultList.Update(model.Header);
                    var h = new TblBuHistoryAction()
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = model.Header.Code,
                        Action = "Cập nhật thông tin",
                    };
                    _dbContext.TblBuHistoryAction.Add(h);

                }
                else
                {
                    model.Header.Status = model.Status.Code == "06" ? "01" : model.Status.Code == "07" ? "01" : model.Status.Code;
                    _dbContext.TblBuCalculateResultList.Update(model.Header);
                    var h = new TblBuHistoryAction()
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = model.Header.Code,
                        Action = model.Status.Code == "02" ? "Trình duyệt" : model.Status.Code == "03" ? "Yêu cầu chỉnh sửa" : model.Status.Code == "04" ? "Phê duyệt" : model.Status.Code == "05" ? "Từ chối" : model.Status.Code == "06" ? "Hủy trình duyệt" : "Hủy phê duyệt",
                        Contents = model.Status.Contents
                    };
                    _dbContext.TblBuHistoryAction.Add(h);
                }

                await _dbContext.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                this.Status = false;
                this.Exception = ex;
            }
        }

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

        public async Task<List<TblMdCustomer>> GetCustomer()
        {
            try
            {
                var data = await _dbContext.TblMdCustomer.Where(x => x.CustomerTypeCode == "BBDO").OrderBy(x => x.Name).ToListAsync();
                return data;
            }
            catch (Exception ex)
            {
                return new List<TblMdCustomer>();
            }
        }

        public async Task<string> ExportExcelPlus(string headerId, CalculateResultModel data)
        {
            try
            {
                //var data = await GetResult(headerId, -1);
                var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "Template", "CoSoTinhMucGiamGia.xlsx");

                using var file = new FileStream(templatePath, FileMode.Open, FileAccess.Read);
                IWorkbook workbook = new XSSFWorkbook(file);

                var styles = new
                {
                    Text = ExcelNPOIExtentions.SetCellStyleText(workbook),
                    TextBold = ExcelNPOIExtentions.SetCellStyleTextBold(workbook),
                    Number = ExcelNPOIExtentions.SetCellStyleNumber(workbook),
                    NumberBold = ExcelNPOIExtentions.SetCellStyleNumberBold(workbook),
                    SignCenter = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true),
                    SignLeft = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, false, false, true),
                    HumanSign = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true, true),
                    Header = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, false)
                };
                #region biến sử dụng
                var header = _dbContext.TblBuCalculateResultList.Where(x => x.Code == headerId).ToList().FirstOrDefault();

                var nguoiKy = _dbContext.TblMdSigner.Where(x => x.Code == header.SignerCode).ToList().FirstOrDefault();

                var Date = header.FDate.ToString("dd/MM/yyyy");
                var Date_2 = header.FDate.ToString("'ngày' dd 'tháng' MM 'năm' yyyy", CultureInfo.InvariantCulture);
                var Date_3 = header.FDate.ToString("dd.MM.yyyy", CultureInfo.InvariantCulture);
                var Hour = header.FDate.ToString("HH'h'00", CultureInfo.InvariantCulture);
                var Time = header.FDate.ToString("HH") + ":00:00";
                string valueHeader = $"Thực hiện: từ {Hour} ngày {Date}";
                string valueHeader_2 = $"Kèm theo Quyết định số: …………............/PLXNA-QĐ ngày {Date}";
                var CVA5 = $"  (Kèm theo Công văn số:                        /PLXNA ngày {header.FDate.Day:D2}/{header.FDate.Month:D2}/{header.FDate.Year} của Công ty Xăng dầu Nghệ An)";
                var QuyetDinhSo = header.QuyetDinhSo;
                #endregion

                #region Dữ liệu gốc
                var startRowdlg_1 = 4;
                ISheet sheetGLG = workbook.GetSheetAt(0);
                #region Thị trường Thành phố Vinh, TX Cửa Lò

                IRow rowHeader_dlg_1 = sheetGLG.GetRow(1);
                ICell header_dlg_1 = rowHeader_dlg_1.GetCell(1) ?? rowHeader_dlg_1.CreateCell(1);
                header_dlg_1.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_1.SetCellValue($"1. Thị trường Thành phố Vinh, TX Cửa Lò (áp dụng từ {Hour} ngày {Date})");


                for (var i = 0; i < data.DLG.Dlg_1.Count(); i++)
                {
                    var dataDlg1 = data.DLG.Dlg_1[i];
                    int rowIndex = startRowdlg_1 + i;

                    //var textStyle = dataDlg1.IsBold ? styles.TextBold : styles.Text;
                    //var numberStyle = dataDlg1.IsBold ? styles.NumberBold : styles.Number;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {
                        ICell cell3 = row.GetCell(1) ?? row.CreateCell(1);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue(dataDlg1.Col1);

                        ICell cell4 = row.GetCell(4) ?? row.CreateCell(4);
                        //cell4.CellStyle = numberStyle; ;
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg1.Col2));

                        ICell cell5 = row.GetCell(5) ?? row.CreateCell(5);
                        //cell5.CellStyle = numberStyle; ;
                        cell5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell5.SetCellValue(Convert.ToDouble(dataDlg1.Col3));

                        ICell cell6 = row.GetCell(6) ?? row.CreateCell(6);
                        //cell6.CellStyle = numberStyle; ;
                        cell6.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell6.SetCellValue( Convert.ToDouble(dataDlg1.Col4));

                        ICell cell7 = row.GetCell(7) ?? row.CreateCell(7);
                        //cell7.CellStyle = numberStyle; ;
                        cell7.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell7.SetCellValue(Convert.ToDouble(dataDlg1.Col5));

                        ICell cell8 = row.GetCell(8) ?? row.CreateCell(8);
                        //cell8.CellStyle = numberStyle; ;
                        cell8.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell8.SetCellValue(Convert.ToDouble(dataDlg1.Col6));

                        ICell cell9 = row.GetCell(9) ?? row.CreateCell(9);
                        //cell9.CellStyle = numberStyle; ;
                        cell9.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell9.SetCellValue(Convert.ToDouble(dataDlg1.Col7));
                    }
                }
                #endregion

                #region Các huyện thị còn lại trên địa bàn Nghệ An + địa bàn tỉnh Hà Tĩnh

                IRow rowHeader_dlg_2 = sheetGLG.GetRow(9);
                ICell header_dlg_2 = rowHeader_dlg_2.GetCell(1) ?? rowHeader_dlg_2.CreateCell(1);
                header_dlg_2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_2.SetCellValue($"2. Các huyện thị còn lại trên địa bàn Nghệ An + địa bàn tỉnh Hà Tĩnh");

                var startRowdlg_2 = 11;
                for (var i = 0; i < data.DLG.Dlg_2.Count(); i++)
                {
                    var dataDlg2 = data.DLG.Dlg_2[i];
                    int rowIndex = startRowdlg_2 + i;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {
                        ICell cell3 = row.GetCell(1) ?? row.CreateCell(1);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue(dataDlg2.Col1);

                        ICell cell4 = row.GetCell(4) ?? row.CreateCell(4);
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg2.Col2));
                    }
                }

                #endregion

                #region BIỂU TỔNG HỢP CÁC CHỈ TIÊU DẦU SÁNG (PT bán lẻ - V2)

                IRow rowHeader_dlg_3 = sheetGLG.GetRow(36);
                ICell header_dlg_3 = rowHeader_dlg_3.GetCell(0) ?? rowHeader_dlg_3.CreateCell(0);
                header_dlg_3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_3.SetCellValue($"Tính từ {Hour} ngày {Date} theo CĐ số {QuyetDinhSo} ngày {Date}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {Date} và theo VCF Hè Thu");

                IRow rowFooter_dlg_3 = sheetGLG.GetRow(46);
                ICell footer_dlg_3 = rowFooter_dlg_3.GetCell(9) ?? rowFooter_dlg_3.CreateCell(9);
                footer_dlg_3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                footer_dlg_3.SetCellValue($"Vinh, {Date_2}");

                if (header.SignerCode == "TongGiamDoc")
                {
                    IRow rowFooter_Ky_dlg_3 = sheetGLG.GetRow(47);
                    ICell footer_Ky_dlg_3 = rowFooter_Ky_dlg_3.GetCell(9) ?? rowFooter_Ky_dlg_3.CreateCell(9);
                    footer_Ky_dlg_3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_3.SetCellValue($"{nguoiKy.Position}");
                }
                else
                {

                    IRow rowFooter_Ky_dlg_3 = sheetGLG.GetRow(47);
                    ICell footer_Ky_dlg_3 = rowFooter_Ky_dlg_3.GetCell(9) ?? rowFooter_Ky_dlg_3.CreateCell(9);
                    footer_Ky_dlg_3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_3.SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");

                    IRow rowFooter_Ky_dlg_4 = sheetGLG.GetRow(48);
                    ICell footer_Ky_dlg_4 = rowFooter_Ky_dlg_4.GetCell(9) ?? rowFooter_Ky_dlg_4.CreateCell(9);
                    footer_Ky_dlg_4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_4.SetCellValue($"{nguoiKy.Position}");

                }

                var startRowdlg_3 = 41;
                for (var i = 0; i < data.DLG.Dlg_3.Count(); i++)
                {
                    var dataDlg3 = data.DLG.Dlg_3[i];
                    int rowIndex = startRowdlg_3 + i;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {
                        ICell cell0 = row.GetCell(0) ?? row.CreateCell(0);
                        cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell0.SetCellValue(dataDlg3.ColA);

                        ICell cell1 = row.GetCell(1) ?? row.CreateCell(1);
                        cell1.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell1.SetCellValue(dataDlg3.ColB);

                        ICell cell2 = row.GetCell(2) ?? row.CreateCell(2);
                        cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell2.SetCellValue(Convert.ToDouble(dataDlg3.Col1));

                        ICell cell3 = row.GetCell(3) ?? row.CreateCell(3);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue(Convert.ToDouble(dataDlg3.Col2));

                        ICell cell4 = row.GetCell(4) ?? row.CreateCell(4);
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg3.Col3));

                        ICell cell5 = row.GetCell(5) ?? row.CreateCell(5);
                        cell5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell5.SetCellValue(Convert.ToDouble(dataDlg3.Col4));

                        ICell cell6 = row.GetCell(6) ?? row.CreateCell(6);
                        cell6.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell6.SetCellValue(Convert.ToDouble(dataDlg3.Col5));

                        ICell cell7 = row.GetCell(7) ?? row.CreateCell(7);
                        cell7.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell7.SetCellValue(Convert.ToDouble(dataDlg3.Col6));

                        ICell cell8 = row.GetCell(8) ?? row.CreateCell(8);
                        cell8.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell8.SetCellValue(Convert.ToDouble(dataDlg3.Col7));

                        ICell cell10 = row.GetCell(10) ?? row.CreateCell(10);
                        cell10.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell10.SetCellValue( Convert.ToDouble(dataDlg3.Col8) );

                        ICell cell11 = row.GetCell(12) ?? row.CreateCell(12);
                        cell11.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell11.SetCellValue(Convert.ToDouble(dataDlg3.Col9));

                        ICell cell12 = row.GetCell(14) ?? row.CreateCell(14);
                        cell12.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell12.SetCellValue( Convert.ToDouble(dataDlg3.Col10));
                    }
                }

                #endregion

                #region BIỂU TỔNG HỢP CÁC CHỈ TIÊU DẦU SÁNG (ngoài bán lẻ)

                IRow rowHeader_dlg_4 = sheetGLG.GetRow(74);
                ICell header_dlg_4 = rowHeader_dlg_4.GetCell(0) ?? rowHeader_dlg_4.CreateCell(0);
                header_dlg_4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_4.SetCellValue($"Tính từ {Hour} ngày {Date} theo CĐ số {QuyetDinhSo} ngày {Date}; QĐ giá bán lẻ số 682/PLX-TGĐ ngày {Date} và theo VCF Hè Thu");

                IRow rowFooter_dlg_4 = sheetGLG.GetRow(91);
                ICell footer_dlg_4 = rowFooter_dlg_4.GetCell(9) ?? rowFooter_dlg_4.CreateCell(9);
                footer_dlg_4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                footer_dlg_4.SetCellValue($"Vinh, {Date_2}");

                if (header.SignerCode == "TongGiamDoc")
                {
                    IRow rowFooter_Ky_dlg_4 = sheetGLG.GetRow(92);
                    ICell footer_Ky_dlg_4 = rowFooter_Ky_dlg_4.GetCell(9) ?? rowFooter_Ky_dlg_4.CreateCell(9);
                    footer_Ky_dlg_4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_4.SetCellValue($"{nguoiKy.Position}");
                }
                else
                {

                    IRow rowFooter_Ky_dlg_4 = sheetGLG.GetRow(92);
                    ICell footer_Ky_dlg_4 = rowFooter_Ky_dlg_4.GetCell(9) ?? rowFooter_Ky_dlg_4.CreateCell(9);
                    footer_Ky_dlg_4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_4.SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");

                    IRow rowFooter_Ky_dlg_5 = sheetGLG.GetRow(93);
                    ICell footer_Ky_dlg_5 = rowFooter_Ky_dlg_5.GetCell(9) ?? rowFooter_Ky_dlg_5.CreateCell(9);
                    footer_Ky_dlg_5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_5.SetCellValue($"{nguoiKy.Position}");

                }
                var startRowdlg_4 = 80;

                for (var i = 0; i < data.DLG.Dlg_4.Count(); i++)
                {
                    var dataDlg4 = data.DLG.Dlg_4[i];

                    if (dataDlg4.Type != null)
                    {
                        int rowIndex = startRowdlg_4 + i;
                        IRow row = sheetGLG.GetRow(rowIndex);
                        if (row != null)
                        {
                            ICell cell0 = row.GetCell(0) ?? row.CreateCell(0);
                            cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell0.SetCellValue(dataDlg4.ColA);

                            ICell cell1 = row.GetCell(1) ?? row.CreateCell(1);
                            cell1.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell1.SetCellValue(dataDlg4.ColB);

                            ICell cell2 = row.GetCell(2) ?? row.CreateCell(2);
                            cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell2.SetCellValue(Convert.ToDouble(dataDlg4.Col1));

                            ICell cell3 = row.GetCell(3) ?? row.CreateCell(3);
                            cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell3.SetCellValue(Convert.ToDouble(dataDlg4.Col2));

                            ICell cell4 = row.GetCell(4) ?? row.CreateCell(4);
                            cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell4.SetCellValue(Convert.ToDouble(dataDlg4.Col3));

                            ICell cell5 = row.GetCell(5) ?? row.CreateCell(5);
                            cell5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell5.SetCellValue(Convert.ToDouble(dataDlg4.Col4));

                            ICell cell6 = row.GetCell(6) ?? row.CreateCell(6);
                            cell6.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell6.SetCellValue( Convert.ToDouble(dataDlg4.Col5));

                            ICell cell7 = row.GetCell(7) ?? row.CreateCell(7);
                            cell7.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell7.SetCellValue(Convert.ToDouble(dataDlg4.Col6) );

                            ICell cell8 = row.GetCell(8) ?? row.CreateCell(8);
                            cell8.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell8.SetCellValue( Convert.ToDouble(dataDlg4.Col7) );

                            ICell cell9 = row.GetCell(9) ?? row.CreateCell(9);
                            cell9.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell9.SetCellValue(Convert.ToDouble(dataDlg4.Col8));

                            ICell cell10 = row.GetCell(10) ?? row.CreateCell(10);
                            cell10.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell10.SetCellValue( Convert.ToDouble(dataDlg4.Col9));

                            ICell cell11 = row.GetCell(11) ?? row.CreateCell(11);
                            cell11.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell11.SetCellValue(Convert.ToDouble(dataDlg4.Col10));

                            ICell cell12 = row.GetCell(12) ?? row.CreateCell(12);
                            cell12.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell12.SetCellValue(Convert.ToDouble(dataDlg4.Col11));

                            ICell cell13 = row.GetCell(13) ?? row.CreateCell(13);
                            cell13.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell13.SetCellValue(Convert.ToDouble(dataDlg4.Col12));

                            ICell cell14 = row.GetCell(14) ?? row.CreateCell(14);
                            cell14.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell14.SetCellValue( Convert.ToDouble(dataDlg4.Col13));

                            ICell cell15 = row.GetCell(15) ?? row.CreateCell(15);
                            cell15.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell15.SetCellValue(Convert.ToDouble(dataDlg4.Col14));

                            ICell cell16 = row.GetCell(16) ?? row.CreateCell(16);
                            cell16.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell16.SetCellValue(Convert.ToDouble(dataDlg4.Col15));

                            ICell cell17 = row.GetCell(17) ?? row.CreateCell(17);
                            cell17.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell17.SetCellValue( Convert.ToDouble(dataDlg4.Col16));
                        }
                    }
                }

                #endregion

                #region BIỂU TÍNH GIÁ XUẤT NỘI DỤNG
                IRow rowHeader_dlg_5 = sheetGLG.GetRow(110);
                ICell header_dlg_5 = rowHeader_dlg_5.GetCell(0) ?? rowHeader_dlg_5.CreateCell(0);
                header_dlg_5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_5.SetCellValue($"Tính từ {Hour} ngày {Date} theo CĐ số {QuyetDinhSo} ngày {Date}; QĐ giá bán lẻ số {QuyetDinhSo} ngày {Date} và theo VCF Hè Thu");

                IRow rowFooter_dlg_5 = sheetGLG.GetRow(122);
                ICell footer_dlg_5 = rowFooter_dlg_5.GetCell(9) ?? rowFooter_dlg_5.CreateCell(9);
                footer_dlg_5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                footer_dlg_5.SetCellValue($"Vinh,  {Date_2}");

                if (header.SignerCode == "TongGiamDoc")
                {
                    IRow rowFooter_Ky_dlg_5 = sheetGLG.GetRow(123);
                    ICell footer_Ky_dlg_5 = rowFooter_Ky_dlg_5.GetCell(9) ?? rowFooter_Ky_dlg_5.CreateCell(9);
                    footer_Ky_dlg_5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_5.SetCellValue($"{nguoiKy.Position}");
                }
                else
                {

                    IRow rowFooter_Ky_dlg_5 = sheetGLG.GetRow(123);
                    ICell footer_Ky_dlg_5 = rowFooter_Ky_dlg_5.GetCell(9) ?? rowFooter_Ky_dlg_5.CreateCell(9);
                    footer_Ky_dlg_5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_Ky_dlg_5.SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");

                    IRow rowFooter_ngKy_dlg_5 = sheetGLG.GetRow(124);
                    ICell footer_ngKy_dlg_5 = rowFooter_ngKy_dlg_5.GetCell(9) ?? rowFooter_ngKy_dlg_5.CreateCell(9);
                    footer_ngKy_dlg_5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                    footer_ngKy_dlg_5.SetCellValue($"{nguoiKy.Position}");

                }

                var startRowdlg_5 = 115;
                for (var i = 0; i < data.DLG.Dlg_5.Count(); i++)
                {
                    var dataDlg5 = data.DLG.Dlg_5[i];
                    int rowIndex = startRowdlg_5 + i;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {
                        ICell cell0 = row.GetCell(0) ?? row.CreateCell(0);
                        cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell0.SetCellValue(dataDlg5.ColA);

                        ICell cell1 = row.GetCell(1) ?? row.CreateCell(1);
                        cell1.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell1.SetCellValue(dataDlg5.ColB);

                        ICell cell2 = row.GetCell(3) ?? row.CreateCell(3);
                        cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell2.SetCellValue( Convert.ToDouble(dataDlg5.Col1));

                        ICell cell3 = row.GetCell(5) ?? row.CreateCell(5);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue( Convert.ToDouble(dataDlg5.Col2));

                        ICell cell4 = row.GetCell(8) ?? row.CreateCell(8);
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg5.Col3));

                        ICell cell5 = row.GetCell(10) ?? row.CreateCell(10);
                        cell5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell5.SetCellValue( Convert.ToDouble(dataDlg5.Col4) );

                        ICell cell6 = row.GetCell(12) ?? row.CreateCell(12);
                        cell6.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell6.SetCellValue( Convert.ToDouble(dataDlg5.Col5));
                    }
                }

                #endregion

                #region Thay đổi giá bán lẻ
                IRow rowHeader_tt = sheetGLG.GetRow(38);
                ICell header_tt = rowHeader_tt.GetCell(21) ?? rowHeader_tt.CreateCell(21);
                header_tt.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_tt.SetCellValue($"Thay đổi giá bán lẻ {Hour} ngày {Date} (Tp Vinh, TX Cửa Lò)");

                ICell header_Vcl = rowHeader_tt.GetCell(24) ?? rowHeader_tt.CreateCell(24);
                header_Vcl.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_Vcl.SetCellValue($"Thay đổi giá bán lẻ {Hour} ngày {Date} (Vùng còn lại)");

                var startRowdlg_TDGBL = 41;
                for (var i = 0; i < data.DLG.Dlg_TDGBL.Count(); i++)
                {
                    var dataDlg_TDGBL = data.DLG.Dlg_TDGBL[i];
                    int rowIndex = startRowdlg_TDGBL + i;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {

                        ICell cell0 = row.GetCell(20) ?? row.CreateCell(20);
                        cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell0.SetCellValue(dataDlg_TDGBL.ColA);

                        ICell cell2 = row.GetCell(21) ?? row.CreateCell(21);
                        cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell2.SetCellValue( Convert.ToDouble(dataDlg_TDGBL.Col1));

                        ICell cell3 = row.GetCell(22) ?? row.CreateCell(22);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue(Convert.ToDouble(dataDlg_TDGBL.Col2));

                        ICell cell4 = row.GetCell(23) ?? row.CreateCell(23);
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg_TDGBL.TangGiam1_2) );

                        ICell cell5 = row.GetCell(24) ?? row.CreateCell(24);
                        cell5.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell5.SetCellValue(Convert.ToDouble(dataDlg_TDGBL.Col3));

                        ICell cell6 = row.GetCell(25) ?? row.CreateCell(25);
                        cell6.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell6.SetCellValue(Convert.ToDouble(dataDlg_TDGBL.Col4) );

                        ICell cell7 = row.GetCell(26) ?? row.CreateCell(26);
                        cell7.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell7.SetCellValue( Convert.ToDouble(dataDlg_TDGBL.TangGiam3_4));
                    }
                }

                #endregion

                #region Thay đổi giá giao phương thức bán lẻ

                var startRowdlg_TdGgptbl = 49;
                for (var i = 0; i < data.DLG.Dlg_TdGgptbl.Count(); i++)
                {
                    var dataDlg_TdGgptbl = data.DLG.Dlg_TdGgptbl[i];
                    int rowIndex = startRowdlg_TdGgptbl + i;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {

                        ICell cell0 = row.GetCell(20) ?? row.CreateCell(20);
                        cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell0.SetCellValue(dataDlg_TdGgptbl.ColA);

                        ICell cell2 = row.GetCell(21) ?? row.CreateCell(21);
                        cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell2.SetCellValue(Convert.ToDouble(dataDlg_TdGgptbl.Col1) );

                        ICell cell3 = row.GetCell(22) ?? row.CreateCell(22);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue(Convert.ToDouble(dataDlg_TdGgptbl.Col2) );

                        ICell cell4 = row.GetCell(23) ?? row.CreateCell(23);
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg_TdGgptbl.TangGiam1_2) );
                    }
                }

                #endregion

                #region So sánh lãi gộp giữa

                int rowIndexTT = 12;
                int rowIndexOther = 17;
                IRow rowHeader_dlg_7 = sheetGLG.GetRow(8);
                ICell header_dlg_7 = rowHeader_dlg_7.GetCell(20) ?? rowHeader_dlg_7.CreateCell(20);
                header_dlg_7.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_7.SetCellValue($"1.Lãi gộp từ {Hour} ngày {Date} và tính theo VCF Hè Thu từ tháng 5 - 10 hàng năm");
                foreach (var dataDlg_Dlg7 in data.DLG.Dlg_7)
                {
                    if (dataDlg_Dlg7.Type == "TT")
                    {
                        IRow row = sheetGLG.GetRow(rowIndexTT);
                        if (row != null)
                        {
                            ICell cell0 = row.GetCell(20) ?? row.CreateCell(20);
                            cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell0.SetCellValue(dataDlg_Dlg7.ColA);

                            ICell cell2 = row.GetCell(21) ?? row.CreateCell(21);
                            cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell2.SetCellValue(Convert.ToDouble(dataDlg_Dlg7.Col1));

                            ICell cell3 = row.GetCell(22) ?? row.CreateCell(22);
                            cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell3.SetCellValue(Convert.ToDouble(dataDlg_Dlg7.Col2));

                            ICell cell4 = row.GetCell(23) ?? row.CreateCell(23);
                            cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell4.SetCellValue(Convert.ToDouble(dataDlg_Dlg7.TangGiam1_2));
                        }
                        rowIndexTT++; // Tăng dòng sau mỗi lần lặp
                    }
                    else if (dataDlg_Dlg7.Type == "OTHER")
                    {
                        IRow row = sheetGLG.GetRow(rowIndexOther);
                        if (row != null)
                        {
                            ICell cell0 = row.GetCell(20) ?? row.CreateCell(20);
                            cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell0.SetCellValue(dataDlg_Dlg7.ColA);

                            ICell cell2 = row.GetCell(21) ?? row.CreateCell(21);
                            cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell2.SetCellValue(Convert.ToDouble(dataDlg_Dlg7.Col1));

                            ICell cell3 = row.GetCell(22) ?? row.CreateCell(22);
                            cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell3.SetCellValue(Convert.ToDouble(dataDlg_Dlg7.Col2));

                            ICell cell4 = row.GetCell(23) ?? row.CreateCell(23);
                            cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                            cell4.SetCellValue( Convert.ToDouble(dataDlg_Dlg7.TangGiam1_2));
                        }
                        rowIndexOther++; // Tăng dòng sau mỗi lần lặp
                    }
                }

                #region So sánh chiết khấu giữa
                IRow rowHeader_dlg_8 = sheetGLG.GetRow(23);
                ICell header_dlg_8 = rowHeader_dlg_8.GetCell(20) ?? rowHeader_dlg_8.CreateCell(20);
                header_dlg_8.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_dlg_8.SetCellValue($"2. Đề xuất mức giảm giá từ {Hour} ngày {Date}");
                var startRowdlg_Dlg8 = 26;
                for (var i = 0; i < data.DLG.Dlg_8.Count(); i++)
                {
                    var dataDlg_Dlg8 = data.DLG.Dlg_8[i];
                    int rowIndex = startRowdlg_Dlg8 + i;
                    IRow row = sheetGLG.GetRow(rowIndex);

                    if (row != null)
                    {
                        ICell cell0 = row.GetCell(20) ?? row.CreateCell(20);
                        cell0.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell0.SetCellValue(dataDlg_Dlg8.ColA);

                        ICell cell2 = row.GetCell(21) ?? row.CreateCell(21);
                        cell2.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell2.SetCellValue( Convert.ToDouble(dataDlg_Dlg8.Col1));

                        ICell cell3 = row.GetCell(22) ?? row.CreateCell(22);
                        cell3.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell3.SetCellValue( Convert.ToDouble(dataDlg_Dlg8.Col2) );

                        ICell cell4 = row.GetCell(23) ?? row.CreateCell(23);
                        cell4.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                        cell4.SetCellValue(Convert.ToDouble(dataDlg_Dlg8.TangGiam1_2));
                    }

                }

                #endregion

                #region Valid
                IRow rowHeader_valid = sheetGLG.GetRow(4);
                ICell header_valid = rowHeader_valid.GetCell(20) ?? rowHeader_valid.CreateCell(20);
                header_valid.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_valid.SetCellValue($"{Date_3}");

                ICell header_valid_time = rowHeader_valid.GetCell(21) ?? rowHeader_valid.CreateCell(21);
                header_valid_time.CellStyle.VerticalAlignment = VerticalAlignment.Center;
                header_valid_time.SetCellValue($"{Time}");
                #endregion

                #endregion

                #endregion 

                #region PT
                var sheetPT = workbook.GetSheetAt(1);
                int startRowPT = 7;

                var rowDatePT = ReportUtilities.CreateRow(ref sheetPT, 1, 1);

                rowDatePT.Cells[0].SetCellValue(valueHeader);
                rowDatePT.Cells[0].CellStyle = styles.Header;

                foreach (var item in data.PT)
                {
                    var textStyle = item.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = item.IsBold ? styles.NumberBold : styles.Number;
                    var rowCur = ReportUtilities.CreateRow(ref sheetPT, startRowPT++, 38);

                    rowCur.Cells[0].CellStyle = textStyle;
                    rowCur.Cells[0].SetCellValue(item.ColA?.ToString());

                    rowCur.Cells[1].SetCellValue(item.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].CellStyle = numberStyle;

                    if (item.IsBold)
                    {
                        rowCur.Cells[2].SetCellValue("");
                    }
                    else
                    {
                        rowCur.Cells[2].SetCellValue(Convert.ToDouble(item.Col1));
                    }

                    for (int lg = 0, col = 3; lg < item.LG.Count; lg++, col++)
                    {
                        rowCur.Cells[col].CellStyle = numberStyle;
                        if (item.IsBold)
                        {
                            rowCur.Cells[col].SetCellValue(Convert.ToDouble(""));
                        }
                        else
                        {
                            rowCur.Cells[col].SetCellValue(Convert.ToDouble(item.LG[lg]));
                        }
                    }
                    rowCur.Cells[11].CellStyle = textStyle;
                    rowCur.Cells[12].CellStyle = textStyle;

                    SetCellValues(rowCur, numberStyle, item, new[] { 8, 9, 10, 11, 12 }, new[] { item.Col3, item.Col4, item.Col5, item.Col6, item.Col7 }, item.IsBold);

                    for (int ggIndex = 0, col = 13; ggIndex < item.GG.Count; ggIndex++, col += 2)
                    {
                        SetCellValues(rowCur, numberStyle, item.GG[ggIndex], new[] { col, col + 1 }, new[] { item.GG[ggIndex].VAT, item.GG[ggIndex].NonVAT }, item.IsBold);
                    }

                    for (int ln = 0, col = 23; ln < item.LN.Count; ln++, col++)
                    {
                        rowCur.Cells[col].CellStyle = numberStyle;
                        if (item.IsBold)
                        {
                            rowCur.Cells[col].SetCellValue("");
                        }
                        else
                        {
                            rowCur.Cells[col].SetCellValue(Convert.ToDouble(item.LN[ln]));
                        }
                    }

                    for (int bvIndex = 0, col = 28; bvIndex < item.BVMT.Count; bvIndex++, col += 2)
                    {
                        SetCellValues(rowCur, numberStyle, item.BVMT[bvIndex], new[] { col, col + 1 }, new[] { item.BVMT[bvIndex].NonVAT, item.BVMT[bvIndex].VAT });
                    }


                }
                var rowSignPT = ReportUtilities.CreateRow(ref sheetPT, startRowPT + 1, 24);

                rowSignPT.Cells[1].SetCellValue("LẬP BIỂU");
                rowSignPT.Cells[1].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);

                rowSignPT.Cells[5].SetCellValue("P. KINH DOANH XD");
                rowSignPT.Cells[5].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, false, false, true);
                sheetPT.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPT + 1, startRowPT + 1, 5, 8));

                rowSignPT.Cells[9].SetCellValue("PHÒNG TCKT");
                rowSignPT.Cells[9].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);
                sheetPT.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPT + 1, startRowPT + 1, 9, 13));

                rowSignPT.Cells[19].SetCellValue("DUYỆT");
                rowSignPT.Cells[19].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);
                sheetPT.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPT + 1, startRowPT + 1, 19, 23));

                #endregion

                #region Export ĐB
                var startRowDB = 7;
                var sheetDB = workbook.GetSheetAt(2);

                var rowDateĐB = ReportUtilities.CreateRow(ref sheetDB, 1, 1);

                rowDateĐB.Cells[0].SetCellValue(valueHeader);
                rowDateĐB.Cells[0].CellStyle = styles.Header;

                foreach (var dataRow in data.DB)
                {
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    var rowCur = ReportUtilities.CreateRow(ref sheetDB, startRowDB++, 43);

                    rowCur.Cells[0].SetCellValue(dataRow.ColA?.ToString());
                    rowCur.Cells[0].CellStyle = textStyle;
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;
                    rowCur.Cells[2].SetCellValue(dataRow.Col1);
                    rowCur.Cells[2].CellStyle = textStyle;

                    rowCur.Cells[3].CellStyle = numberStyle;
                    if (dataRow.IsBold)
                    {
                        rowCur.Cells[3].SetCellValue("");
                    }
                    else
                    {
                        rowCur.Cells[3].SetCellValue(Convert.ToDouble(dataRow.Col2));
                    }

                    for (int lg = 0, col = 4; lg < dataRow.LG.Count; lg++, col++)
                    {
                        rowCur.Cells[col].CellStyle = numberStyle;
                        rowCur.Cells[col].SetCellValue(Convert.ToDouble(dataRow.LG[lg]));
                    }

                    SetCellValues(rowCur, numberStyle, dataRow, Enumerable.Range(9, 8).ToArray(),
                        new[] { dataRow.Col3, dataRow.Col4, dataRow.Col5, dataRow.Col6, dataRow.Col7, dataRow.Col8, dataRow.Col9, dataRow.Col10 }, dataRow.IsBold);

                    for (int gg = 0, col = 17; gg < dataRow.GG.Count; gg++, col += 2)
                    {
                        SetCellValues(rowCur, numberStyle, dataRow.GG[gg], new[] { col, col + 1 }, new[] { dataRow.GG[gg].VAT, dataRow.GG[gg].NonVAT }, dataRow.IsBold);
                    }

                    for (int ln = 0, col = 28; ln < dataRow.LN.Count; ln++, col++)
                    {
                        rowCur.Cells[col].CellStyle = numberStyle;
                        rowCur.Cells[col].SetCellValue(Convert.ToDouble(dataRow.LN[ln]));
                    }

                    for (int bv = 0, col = 33; bv < dataRow.BVMT.Count; bv++, col += 2)
                    {
                        SetCellValues(rowCur, numberStyle, dataRow.BVMT[bv], new[] { col, col + 1 }, new[] { dataRow.BVMT[bv].NonVAT, dataRow.BVMT[bv].VAT }, dataRow.IsBold);
                    }
                }
                var rowSignDB = ReportUtilities.CreateRow(ref sheetDB, startRowDB, 28);

                rowSignDB.Cells[1].SetCellValue("LẬP BIỂU");
                rowSignDB.Cells[1].CellStyle = styles.SignCenter;

                rowSignDB.Cells[6].SetCellValue("P. KINH DOANH XD");
                rowSignDB.Cells[6].CellStyle = styles.SignLeft;
                sheetDB.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowDB, startRowDB, 6, 9));

                rowSignDB.Cells[15].SetCellValue("KẾ  TOÁN TRƯỞNG");
                rowSignDB.Cells[15].CellStyle = styles.SignCenter;
                sheetDB.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowDB, startRowDB, 15, 17));

                rowSignDB.Cells[27].SetCellValue("DUYỆT");
                rowSignDB.Cells[27].CellStyle = styles.SignLeft;
                #endregion

                #region Export FOB
                var startRowFOB = 7;
                var sheetFOB = workbook.GetSheetAt(3);

                var rowDateFOB = ReportUtilities.CreateRow(ref sheetFOB, 1, 1);

                rowDateFOB.Cells[0].SetCellValue(valueHeader);
                rowDateFOB.Cells[0].CellStyle = styles.Header;

                foreach (var dataRow in data.FOB)
                {
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    var rowCur = ReportUtilities.CreateRow(ref sheetFOB, startRowFOB++, 41);

                    rowCur.Cells[0].SetCellValue(dataRow.ColA?.ToString());
                    rowCur.Cells[0].CellStyle = textStyle;
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    for (int lg = 0, col = 2; lg < dataRow.LG.Count; lg++, col++)
                    {
                        rowCur.Cells[col].CellStyle = numberStyle;
                        rowCur.Cells[col].SetCellValue(Convert.ToDouble(dataRow.LG[lg]));
                    }

                    SetCellValues(rowCur, numberStyle, dataRow, Enumerable.Range(7, 7).ToArray(),
                        new[] { dataRow.Col1, dataRow.Col2, dataRow.Col3, dataRow.Col4, dataRow.Col5, dataRow.Col6, dataRow.Col7 }, dataRow.IsBold);

                    for (int ggIndex = 0, col = 14; ggIndex < dataRow.GG.Count; ggIndex++, col += 2)
                    {
                        SetCellValues(rowCur, numberStyle, dataRow.GG[ggIndex], new[] { col, col + 1 }, new[] { dataRow.GG[ggIndex].VAT, dataRow.GG[ggIndex].NonVAT });
                    }

                    rowCur.Cells[22].SetCellValue(dataRow.Col8 == 0 ? 0 : Convert.ToDouble(dataRow.Col8));
                    rowCur.Cells[22].CellStyle = textStyle;

                    for (int ln = 0, col = 25; ln < dataRow.LN.Count; ln++, col++)
                    {
                        rowCur.Cells[col].CellStyle = numberStyle;
                        rowCur.Cells[col].SetCellValue(Convert.ToDouble(dataRow.LN[ln]));
                    }

                    for (int bvIndex = 0, col = 30; bvIndex < dataRow.BVMT.Count; bvIndex++, col += 2)
                    {
                        SetCellValues(rowCur, numberStyle, dataRow.BVMT[bvIndex], new[] { col, col + 1 }, new[] { dataRow.BVMT[bvIndex].NonVAT, dataRow.BVMT[bvIndex].VAT });
                    }
                }
                var rowSignFOB = ReportUtilities.CreateRow(ref sheetFOB, startRowFOB + 1, 24);

                rowSignFOB.Cells[1].SetCellValue("LẬP BIỂU");
                rowSignFOB.Cells[1].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);

                rowSignFOB.Cells[5].SetCellValue("P. KINH DOANH XD");
                rowSignFOB.Cells[5].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, false, false, true);
                sheetFOB.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowFOB + 1, startRowFOB + 1, 5, 8));

                rowSignFOB.Cells[13].SetCellValue("KẾ  TOÁN TRƯỞNG");
                rowSignFOB.Cells[13].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, false, false, true);
                sheetFOB.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowFOB + 1, startRowFOB + 1, 13, 15));

                rowSignFOB.Cells[20].SetCellValue("DUYỆT");
                rowSignFOB.Cells[20].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);
                sheetFOB.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowFOB + 1, startRowFOB + 1, 20, 24));
                #endregion

                #region Export PT09

                var startRowPT09 = 7;
                ISheet sheetPT09 = workbook.GetSheetAt(4);

                var rowDatePT09 = ReportUtilities.CreateRow(ref sheetPT09, 1, 1);

                rowDatePT09.Cells[0].SetCellValue(valueHeader);
                rowDatePT09.Cells[0].CellStyle = styles.Header;

                for (var i = 0; i < data.PT09.Count(); i++)
                {
                    var dataRow = data.PT09[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;

                    IRow rowCur = ReportUtilities.CreateRow(ref sheetPT09, startRowPT09++, 39);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    var iLG = 0;
                    for (var lg = iLG; lg < dataRow.LG.Count(); lg++)
                    {
                        rowCur.Cells[2 + lg].CellStyle = numberStyle;
                        rowCur.Cells[2 + lg].SetCellValue(dataRow.LG[lg] == 0 ? 0 : Convert.ToDouble(dataRow.LG[lg]));
                    }

                    rowCur.Cells[7].CellStyle = numberStyle; ;
                    rowCur.Cells[7].SetCellValue(dataRow.Col3 == 0 ? 0 : Convert.ToDouble(dataRow.Col3));

                    rowCur.Cells[8].CellStyle = numberStyle; ;
                    rowCur.Cells[8].SetCellValue(dataRow.Col4 == 0 ? 0 : Convert.ToDouble(dataRow.Col4));

                    rowCur.Cells[9].CellStyle = numberStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col5 == 0 ? 0 : Convert.ToDouble(dataRow.Col5));

                    rowCur.Cells[10].CellStyle = numberStyle; ;
                    rowCur.Cells[10].SetCellValue(dataRow.Col6 == 0 ? 0 : Convert.ToDouble(dataRow.Col6));

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col7 == 0 ? 0 : Convert.ToDouble(dataRow.Col7));

                    rowCur.Cells[12].CellStyle = numberStyle; ;
                    rowCur.Cells[12].SetCellValue(dataRow.Col8 == 0 ? 0 : Convert.ToDouble(dataRow.Col8));


                    var iGG = 0;
                    for (var gg = 0; gg < dataRow.GG.Count(); gg++)
                    {
                        rowCur.Cells[13 + iGG].CellStyle = numberStyle; ;
                        rowCur.Cells[13 + iGG].SetCellValue(dataRow.GG[gg].VAT == 0 ? 0 : Convert.ToDouble(dataRow.GG[gg].VAT));
                        rowCur.Cells[14 + iGG].CellStyle = numberStyle; ;
                        rowCur.Cells[14 + iGG].SetCellValue(dataRow.GG[gg].NonVAT == 0 ? 0 : Convert.ToDouble(dataRow.GG[gg].NonVAT));
                        iGG += 2;
                    }

                    rowCur.Cells[23].CellStyle = numberStyle; ;
                    rowCur.Cells[23].SetCellValue(dataRow.Col18 == 0 ? 0 : Convert.ToDouble(dataRow.Col18));

                    var iLN = 0;
                    for (var ln = iLN; ln < dataRow.LG.Count(); ln++)
                    {
                        rowCur.Cells[24 + ln].CellStyle = numberStyle; ;
                        rowCur.Cells[24 + ln].SetCellValue(dataRow.LN[ln] == 0 ? 0 : Convert.ToDouble(dataRow.LN[ln]));
                    }

                    var iBV = 0;
                    for (var gg = 0; gg < dataRow.BVMT.Count(); gg++)
                    {
                        rowCur.Cells[29 + iBV].CellStyle = numberStyle; ;
                        rowCur.Cells[29 + iBV].SetCellValue(dataRow.BVMT[gg].NonVAT == 0 ? 0 : Convert.ToDouble(dataRow.BVMT[gg].NonVAT));
                        rowCur.Cells[30 + iBV].CellStyle = numberStyle; ;
                        rowCur.Cells[30 + iBV].SetCellValue(dataRow.BVMT[gg].VAT == 0 ? 0 : Convert.ToDouble(dataRow.BVMT[gg].VAT));
                        iBV += 2;
                    }
                }

                var rowSignPT09 = ReportUtilities.CreateRow(ref sheetPT09, startRowPT09 + 1, 24);

                rowSignPT09.Cells[1].SetCellValue("LẬP BIỂU");
                rowSignPT09.Cells[1].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);

                rowSignPT09.Cells[5].SetCellValue("P. KINH DOANH XD");
                rowSignPT09.Cells[5].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, false, false, true);
                sheetPT09.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPT09 + 1, startRowPT09 + 1, 5, 8));

                rowSignPT09.Cells[9].SetCellValue("KẾ  TOÁN TRƯỞNG");
                rowSignPT09.Cells[9].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);
                sheetPT09.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPT09 + 1, startRowPT09 + 1, 9, 16));

                rowSignPT09.Cells[19].SetCellValue("DUYỆT");
                rowSignPT09.Cells[19].CellStyle = ExcelNPOIExtentions.SetCellStyleTextSign(workbook, true, false, true);
                sheetPT09.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPT09 + 1, startRowPT09 + 1, 19, 23));
                #endregion

                #region Export BB ĐO

                var startRowBBDO = 9;
                ISheet sheetBBDO = workbook.GetSheetAt(5);

                var rowDateBBDO = ReportUtilities.CreateRow(ref sheetBBDO, 3, 1);

                rowDateBBDO.Cells[0].SetCellValue(valueHeader);
                rowDateBBDO.Cells[0].CellStyle = styles.Header;

                var rowQdsBBDO = ReportUtilities.CreateRow(ref sheetBBDO, 4, 1);

                rowQdsBBDO.Cells[0].SetCellValue(valueHeader_2);
                rowQdsBBDO.Cells[0].CellStyle = styles.Header;

                for (var i = 0; i < data.BBDO.Count(); i++)
                {
                    var dataRow = data.BBDO[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;

                    IRow rowCur = ReportUtilities.CreateRow(ref sheetBBDO, startRowBBDO++, 23);
                    rowCur.Cells[0].CellStyle = numberStyle; ;
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);

                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].SetCellValue(dataRow.ColC);
                    rowCur.Cells[2].CellStyle = textStyle;

                    rowCur.Cells[3].SetCellValue(dataRow.ColD);
                    rowCur.Cells[3].CellStyle = textStyle;

                    rowCur.Cells[4].CellStyle = numberStyle; ;
                    rowCur.Cells[4].SetCellValue(dataRow.Col1);

                    rowCur.Cells[5].CellStyle = numberStyle; ;
                    rowCur.Cells[5].SetCellValue(dataRow.Col2);

                    rowCur.Cells[6].CellStyle = numberStyle; ;
                    rowCur.Cells[6].SetCellValue(dataRow.Col3);

                    rowCur.Cells[7].CellStyle = textStyle;
                    rowCur.Cells[7].SetCellValue(dataRow.Col4);

                    rowCur.Cells[8].CellStyle = textStyle;
                    rowCur.Cells[8].SetCellValue(dataRow.Col5);


                    rowCur.Cells[9].CellStyle = numberStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col6 == 0 ? 0 : Convert.ToDouble(dataRow.Col6));

                    rowCur.Cells[10].CellStyle = numberStyle; ;
                    rowCur.Cells[10].SetCellValue(dataRow.Col7 == 0 ? 0 : Convert.ToDouble(dataRow.Col7));

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col8 == 0 ? 0 : Convert.ToDouble(dataRow.Col8));

                    rowCur.Cells[12].CellStyle = numberStyle; ;
                    rowCur.Cells[12].SetCellValue(dataRow.Col9 == 0 ? 0 : Convert.ToDouble(dataRow.Col9));

                    rowCur.Cells[13].CellStyle = numberStyle; ;
                    rowCur.Cells[13].SetCellValue(dataRow.Col10 == 0 ? 0 : Convert.ToDouble(dataRow.Col10));

                    rowCur.Cells[14].CellStyle = numberStyle; ;
                    rowCur.Cells[14].SetCellValue(dataRow.Col11 == 0 ? 0 : Convert.ToDouble(dataRow.Col11));

                    rowCur.Cells[15].CellStyle = numberStyle; ;
                    rowCur.Cells[15].SetCellValue(dataRow.Col12 == 0 ? 0 : Convert.ToDouble(dataRow.Col12));

                    rowCur.Cells[16].CellStyle = numberStyle; ;
                    rowCur.Cells[16].SetCellValue(dataRow.Col13 == 0 ? 0 : Convert.ToDouble(dataRow.Col13));

                    rowCur.Cells[17].CellStyle = numberStyle; ;
                    rowCur.Cells[17].SetCellValue(dataRow.Col14 == 0 ? 0 : Convert.ToDouble(dataRow.Col14));

                    rowCur.Cells[18].CellStyle = numberStyle; ;
                    rowCur.Cells[18].SetCellValue(dataRow.Col15 == 0 ? 0 : Convert.ToDouble(dataRow.Col15));

                    rowCur.Cells[19].CellStyle = numberStyle; ;
                    rowCur.Cells[19].SetCellValue(dataRow.Col16 == 0 ? 0 : Convert.ToDouble(dataRow.Col16));

                    rowCur.Cells[20].CellStyle = numberStyle; ;
                    rowCur.Cells[20].SetCellValue(dataRow.Col17 == 0 ? 0 : Convert.ToDouble(dataRow.Col17));

                    rowCur.Cells[21].CellStyle = numberStyle; ;
                    rowCur.Cells[21].SetCellValue(dataRow.Col18 == 0 ? 0 : Convert.ToDouble(dataRow.Col18));

                    rowCur.Cells[22].CellStyle = numberStyle; ;
                    rowCur.Cells[22].SetCellValue(dataRow.Col19 == 0 ? 0 : Convert.ToDouble(dataRow.Col19));

                }
                var rowSignBBDO = ReportUtilities.CreateRow(ref sheetBBDO, startRowBBDO + 1, 21);

                rowSignBBDO.Cells[0].SetCellValue("                LẬP BIỂU                                   P.KDXD                                    PHÒNG TCKT");
                rowSignBBDO.Cells[0].CellStyle = styles.SignLeft;

                sheetBBDO.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowBBDO + 1, startRowBBDO + 1, 4, 22));
                if (header.SignerCode == "TongGiamDoc")
                {
                    rowSignBBDO.Cells[4].SetCellValue($"{nguoiKy.Position}");
                    rowSignBBDO.Cells[4].CellStyle = styles.HumanSign;
                    var rowSignBBDO_SignName = ReportUtilities.CreateRow(ref sheetBBDO, startRowBBDO + 6, 25);
                    sheetBBDO.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowBBDO + 6, startRowBBDO + 6, 4, 22));
                    rowSignBBDO_SignName.Cells[4].SetCellValue($"{nguoiKy.Name}");
                    rowSignBBDO_SignName.Cells[4].CellStyle = styles.HumanSign;
                }
                else
                {
                    rowSignBBDO.Cells[4].SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");
                    rowSignBBDO.Cells[4].CellStyle = styles.HumanSign;

                    var rowSignPositionBBDO = ReportUtilities.CreateRow(ref sheetBBDO, startRowBBDO + 2, 22);
                    sheetBBDO.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowBBDO + 2, startRowBBDO + 2, 4, 22));
                    rowSignPositionBBDO.Cells[4].SetCellValue($"{nguoiKy.Position}");
                    rowSignPositionBBDO.Cells[4].CellStyle = styles.HumanSign;

                    var rowSignBBDO_SignName = ReportUtilities.CreateRow(ref sheetBBDO, startRowBBDO + 7, 22);
                    sheetBBDO.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowBBDO + 7, startRowBBDO + 7, 4, 22));
                    rowSignBBDO_SignName.Cells[4].SetCellValue($"{nguoiKy.Name}");
                    rowSignBBDO_SignName.Cells[4].CellStyle = styles.HumanSign;
                }
                #endregion

                #region Export BB FO

                var startRowBBFO = 11;
                ISheet sheetBBFO = workbook.GetSheetAt(6);

                for (var i = 0; i < data.BBFO.Count(); i++)
                {
                    var dataRow = data.BBFO[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetBBFO, startRowBBFO++, 14);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[2].SetCellValue(dataRow.ColC);


                    rowCur.Cells[4].CellStyle = numberStyle; ;
                    rowCur.Cells[4].SetCellValue(dataRow.Col1 == 0 ? 0 : Convert.ToDouble(dataRow.Col1));

                    rowCur.Cells[5].CellStyle = numberStyle; ;
                    rowCur.Cells[5].SetCellValue(dataRow.Col2 == 0 ? 0 : Convert.ToDouble(dataRow.Col2));

                    rowCur.Cells[6].CellStyle = numberStyle; ;
                    rowCur.Cells[6].SetCellValue(dataRow.Col3 == 0 ? 0 : Convert.ToDouble(dataRow.Col3));

                    rowCur.Cells[7].CellStyle = numberStyle; ;
                    rowCur.Cells[7].SetCellValue(dataRow.Col4 == 0 ? 0 : Convert.ToDouble(dataRow.Col4));

                    rowCur.Cells[8].CellStyle = numberStyle; ;
                    rowCur.Cells[8].SetCellValue(dataRow.Col5 == 0 ? 0 : Convert.ToDouble(dataRow.Col5));

                    rowCur.Cells[9].CellStyle = numberStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col6 == 0 ? 0 : Convert.ToDouble(dataRow.Col6));

                    rowCur.Cells[10].CellStyle = numberStyle; ;
                    rowCur.Cells[10].SetCellValue(dataRow.Col7 == 0 ? 0 : Convert.ToDouble(dataRow.Col7));

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col8 == 0 ? 0 : Convert.ToDouble(dataRow.Col8));

                    rowCur.Cells[12].CellStyle = numberStyle; ;
                    rowCur.Cells[12].SetCellValue(dataRow.Col9 == 0 ? 0 : Convert.ToDouble(dataRow.Col9));

                    rowCur.Cells[13].CellStyle = numberStyle; ;
                    rowCur.Cells[13].SetCellValue(dataRow.Col10 == 0 ? 0 : Convert.ToDouble(dataRow.Col10));
                }

                #endregion

                #region Export PL1
                var startRowPL1 = 8;
                ISheet sheetPL1 = workbook.GetSheetAt(7);

                var rowDatePL1 = ReportUtilities.CreateRow(ref sheetPL1, 2, 1);

                rowDatePL1.Cells[0].SetCellValue(valueHeader);
                rowDatePL1.Cells[0].CellStyle = styles.HumanSign;

                for (var i = 0; i < data.PL1.Count(); i++)
                {
                    var dataRow = data.PL1[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetPL1, startRowPL1++, 7);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    var iGG = 0;
                    for (var gg = 0; gg < dataRow.GG.Count(); gg++)
                    {
                        rowCur.Cells[2 + iGG].CellStyle = numberStyle; ;
                        rowCur.Cells[2 + iGG].SetCellValue(dataRow.GG[gg] == 0 ? 0 : Convert.ToDouble(dataRow.GG[gg]));
                        iGG += 1;
                    }

                }
                var rowSignPL1 = ReportUtilities.CreateRow(ref sheetPL1, startRowPL1 + 1, 6);

                rowSignPL1.Cells[0].SetCellValue("LẬP BIỂU            P.KDXD              P.TCKT          ");
                rowSignPL1.Cells[0].CellStyle = styles.SignLeft;
                sheetPL1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL1 + 1, startRowPL1 + 1, 0, 2));
                sheetPL1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL1 + 1, startRowPL1 + 1, 3, 5));
                if (header.SignerCode == "TongGiamDoc")
                {
                    rowSignPL1.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPL1.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPL1_SignName = ReportUtilities.CreateRow(ref sheetPL1, startRowPL1 + 6, 25);
                    sheetPL1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL1 + 6, startRowPL1 + 6, 3, 5));
                    rowSignPL1_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL1_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                else
                {
                    rowSignPL1.Cells[3].SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");
                    rowSignPL1.Cells[3].CellStyle = styles.HumanSign;

                    var rowSignPositionPL1 = ReportUtilities.CreateRow(ref sheetPL1, startRowPL1 + 2, 5);
                    sheetPL1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL1 + 2, startRowPL1 + 2, 3, 5));
                    rowSignPositionPL1.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPositionPL1.Cells[3].CellStyle = styles.HumanSign;

                    var rowSignPL1_SignName = ReportUtilities.CreateRow(ref sheetPL1, startRowPL1 + 7, 6);
                    sheetPL1.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL1 + 7, startRowPL1 + 7, 3, 5));
                    rowSignPL1_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL1_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                #endregion

                #region Export PL2
                var startRowPL2 = 7;
                ISheet sheetPL2 = workbook.GetSheetAt(8);

                var rowDatePL2 = ReportUtilities.CreateRow(ref sheetPL2, 2, 1);

                rowDatePL2.Cells[0].SetCellValue(valueHeader);
                rowDatePL2.Cells[0].CellStyle = styles.HumanSign;

                for (var i = 0; i < data.PL2.Count(); i++)
                {
                    var dataRow = data.PL2[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetPL2, startRowPL2++, 7);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    var iGG = 0;
                    for (var gg = 0; gg < dataRow.GG.Count(); gg++)
                    {
                        rowCur.Cells[2 + iGG].CellStyle = numberStyle; ;
                        rowCur.Cells[2 + iGG].SetCellValue(dataRow.GG[gg] == 0 ? 0 : Convert.ToDouble(dataRow.GG[gg]));
                        iGG += 1;
                    }
                }
                var rowSignPL2 = ReportUtilities.CreateRow(ref sheetPL2, startRowPL2 + 1, 6);

                rowSignPL2.Cells[0].SetCellValue("     LẬP BIỂU            P.KDXD         P.TCKT        ");
                rowSignPL2.Cells[0].CellStyle = styles.SignLeft;
                sheetPL2.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL2 + 1, startRowPL2 + 1, 0, 2));
                sheetPL2.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL2 + 1, startRowPL2 + 1, 3, 5));
                if (header.SignerCode == "TongGiamDoc")
                {
                    rowSignPL2.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPL2.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPL2_SignName = ReportUtilities.CreateRow(ref sheetPL2, startRowPL2 + 6, 25);
                    sheetPL2.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL2 + 6, startRowPL2 + 6, 3, 5));
                    rowSignPL2_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL2_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                else
                {
                    rowSignPL2.Cells[3].SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");
                    rowSignPL2.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPositionPL2 = ReportUtilities.CreateRow(ref sheetPL2, startRowPL2 + 2, 5);
                    sheetPL2.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL2 + 2, startRowPL2 + 2, 3, 5));
                    rowSignPositionPL2.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPositionPL2.Cells[3].CellStyle = styles.HumanSign;

                    var rowSignPL2_SignName = ReportUtilities.CreateRow(ref sheetPL2, startRowPL2 + 7, 6);
                    sheetPL2.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL2 + 7, startRowPL2 + 7, 3, 5));
                    rowSignPL2_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL2_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                #endregion

                #region Export PL3

                var startRowPL3 = 7;
                ISheet sheetPL3 = workbook.GetSheetAt(9);

                var rowDatePL3 = ReportUtilities.CreateRow(ref sheetPL3, 2, 1);

                rowDatePL3.Cells[0].SetCellValue(valueHeader);
                rowDatePL3.Cells[0].CellStyle = styles.HumanSign;

                for (var i = 0; i < data.PL3.Count(); i++)
                {
                    var dataRow = data.PL3[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetPL3, startRowPL3++, 7);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    var iGG = 0;
                    for (var gg = 0; gg < dataRow.GG.Count(); gg++)
                    {
                        rowCur.Cells[2 + iGG].CellStyle = numberStyle; ;
                        rowCur.Cells[2 + iGG].SetCellValue(dataRow.GG[gg] == 0 ? 0 : Convert.ToDouble(dataRow.GG[gg]));
                        iGG += 1;
                    }
                }
                var rowSignPL3 = ReportUtilities.CreateRow(ref sheetPL3, startRowPL3 + 1, 6);

                rowSignPL3.Cells[0].SetCellValue("     LẬP BIỂU            P.KDXD                  P.TCKT   ");
                rowSignPL3.Cells[0].CellStyle = styles.SignLeft;
                sheetPL3.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL3 + 1, startRowPL3 + 1, 0, 2));
                sheetPL3.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL3 + 1, startRowPL3 + 1, 3, 5));
                if (header.SignerCode == "TongGiamDoc")
                {
                    rowSignPL3.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPL3.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPL3_SignName = ReportUtilities.CreateRow(ref sheetPL3, startRowPL3 + 6, 25);
                    sheetPL3.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL3 + 6, startRowPL3 + 6, 3, 5));
                    rowSignPL3_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL3_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                else
                {
                    rowSignPL3.Cells[3].SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");
                    rowSignPL3.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPositionPL3 = ReportUtilities.CreateRow(ref sheetPL3, startRowPL3 + 2, 5);
                    sheetPL3.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL3 + 2, startRowPL3 + 2, 3, 5));
                    rowSignPositionPL3.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPositionPL3.Cells[3].CellStyle = styles.HumanSign;

                    var rowSignPL3_SignName = ReportUtilities.CreateRow(ref sheetPL3, startRowPL3 + 7, 6);
                    sheetPL3.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL3 + 7, startRowPL3 + 7, 3, 5));
                    rowSignPL3_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL3_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                #endregion

                #region Export PL4

                var startRowPL4 = 8;
                ISheet sheetPL4 = workbook.GetSheetAt(10);

                var rowDatePL4 = ReportUtilities.CreateRow(ref sheetPL4, 3, 1);

                rowDatePL4.Cells[2].SetCellValue(valueHeader);
                rowDatePL4.Cells[2].CellStyle = styles.HumanSign;

                for (var i = 0; i < data.PL4.Count(); i++)
                {
                    var dataRow = data.PL4[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetPL4, startRowPL4++, 7);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    var iGG = 0;
                    for (var gg = 0; gg < dataRow.GG.Count(); gg++)
                    {
                        rowCur.Cells[2 + iGG].CellStyle = numberStyle; ;
                        rowCur.Cells[2 + iGG].SetCellValue(dataRow.GG[gg] == 0 ? 0 : Convert.ToDouble(dataRow.GG[gg]));
                        iGG += 1;
                    }
                }
                var rowSignPL4 = ReportUtilities.CreateRow(ref sheetPL4, startRowPL4 + 1, 6);

                rowSignPL4.Cells[0].SetCellValue("     LẬP BIỂU            P.KDXD                  P.TCKT          ");
                rowSignPL4.Cells[0].CellStyle = styles.SignLeft;
                sheetPL4.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL4 + 1, startRowPL4 + 1, 0, 2));
                sheetPL4.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL4 + 1, startRowPL4 + 1, 3, 5));
                if (header.SignerCode == "TongGiamDoc")
                {
                    rowSignPL4.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPL4.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPL4_SignName = ReportUtilities.CreateRow(ref sheetPL4, startRowPL4 + 6, 25);
                    sheetPL4.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL4 + 6, startRowPL4 + 6, 3, 5));
                    rowSignPL4_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL4_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                else
                {
                    rowSignPL4.Cells[3].SetCellValue("KT.CHỦ TỊCH KIÊM GIÁM ĐỐC");
                    rowSignPL4.Cells[3].CellStyle = styles.HumanSign;
                    var rowSignPositionPL4 = ReportUtilities.CreateRow(ref sheetPL4, startRowPL4 + 2, 5);
                    sheetPL4.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL4 + 2, startRowPL4 + 2, 3, 5));
                    rowSignPositionPL4.Cells[3].SetCellValue($"{nguoiKy.Position}");
                    rowSignPositionPL4.Cells[3].CellStyle = styles.HumanSign;

                    var rowSignPL4_SignName = ReportUtilities.CreateRow(ref sheetPL4, startRowPL4 + 7, 6);
                    sheetPL4.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(startRowPL4 + 7, startRowPL4 + 7, 3, 5));
                    rowSignPL4_SignName.Cells[3].SetCellValue($"{nguoiKy.Name}");
                    rowSignPL4_SignName.Cells[3].CellStyle = styles.HumanSign;
                }
                #endregion

                #region Export VK11-PT (11s)

                var startRowVK11PT = 2;
                ISheet sheetVK11PT = workbook.GetSheetAt(11);
                for (var i = 0; i < data.VK11PT.Count(); i++)
                {
                    var dataRow = data.VK11PT[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetVK11PT, startRowVK11PT++, 20);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].CellStyle = textStyle;
                    rowCur.Cells[2].SetCellValue(dataRow.Col1);

                    rowCur.Cells[3].CellStyle = numberStyle;

                    rowCur.Cells[4].CellStyle = numberStyle;

                    rowCur.Cells[10].CellStyle = numberStyle;

                    rowCur.Cells[12].CellStyle = numberStyle;

                    if (dataRow.IsBold)
                    {
                        rowCur.Cells[3].SetCellValue("");
                        rowCur.Cells[4].SetCellValue("");
                        rowCur.Cells[10].SetCellValue("");
                        rowCur.Cells[12].SetCellValue("");
                    }
                    else
                    {
                        rowCur.Cells[3].SetCellValue(dataRow.Col2 == 0 ? 0 : Convert.ToDouble(dataRow.Col2));
                        rowCur.Cells[4].SetCellValue(dataRow.Col3 == 0 ? 0 : Convert.ToDouble(dataRow.Col3));
                        rowCur.Cells[10].SetCellValue(dataRow.Col9 == 0 ? 0 : Convert.ToDouble(dataRow.Col9));
                        rowCur.Cells[12].SetCellValue(dataRow.Col11 == 0 ? 0 : Convert.ToDouble(dataRow.Col11));
                    }

                    rowCur.Cells[5].CellStyle = numberStyle; ;
                    rowCur.Cells[5].SetCellValue(dataRow.Col4);

                    rowCur.Cells[6].CellStyle = numberStyle; ;
                    rowCur.Cells[6].SetCellValue(dataRow.Col5);

                    rowCur.Cells[7].CellStyle = numberStyle; ;
                    rowCur.Cells[7].SetCellValue(dataRow.Col6);

                    rowCur.Cells[8].CellStyle = textStyle; ;
                    rowCur.Cells[8].SetCellValue(dataRow.Col7);

                    rowCur.Cells[9].CellStyle = numberStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col8);

                    rowCur.Cells[11].CellStyle = textStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col10);

                    rowCur.Cells[13].CellStyle = textStyle; ;
                    rowCur.Cells[13].SetCellValue(dataRow.Col12);

                    rowCur.Cells[14].CellStyle = textStyle; ;
                    rowCur.Cells[14].SetCellValue(dataRow.Col13);

                    rowCur.Cells[15].CellStyle = numberStyle; ;
                    rowCur.Cells[15].SetCellValue(dataRow.Col14);

                    rowCur.Cells[16].CellStyle = numberStyle; ;
                    rowCur.Cells[16].SetCellValue(dataRow.Col15);

                    rowCur.Cells[17].CellStyle = numberStyle; ;
                    rowCur.Cells[17].SetCellValue(dataRow.Col16);

                    rowCur.Cells[18].SetCellValue(dataRow.Col17);
                    rowCur.Cells[19].SetCellValue(dataRow.Col18);
                }

                #endregion

                #region Export VK11-DB (3.5s)

                var startRowVK11DB = 2;
                ISheet sheetVK11DB = workbook.GetSheetAt(12);

                for (var i = 0; i < data.VK11DB.Count(); i++)
                {
                    var dataRow = data.VK11DB[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetVK11DB, startRowVK11DB++, 21);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);

                    rowCur.Cells[1].CellStyle = textStyle;
                    rowCur.Cells[1].SetCellValue(dataRow.ColB);

                    rowCur.Cells[2].CellStyle = textStyle;
                    rowCur.Cells[2].SetCellValue(dataRow.ColC);

                    rowCur.Cells[3].CellStyle = textStyle;
                    rowCur.Cells[3].SetCellValue(dataRow.Col1);

                    rowCur.Cells[4].CellStyle = numberStyle; ;
                    rowCur.Cells[4].SetCellValue(dataRow.Col2 == 0 ? 0 : Convert.ToDouble(dataRow.Col2));

                    rowCur.Cells[5].CellStyle = numberStyle; ;
                    rowCur.Cells[5].SetCellValue(dataRow.Col3 == 0 ? 0 : Convert.ToDouble(dataRow.Col3));

                    rowCur.Cells[6].CellStyle = numberStyle; ;
                    rowCur.Cells[6].SetCellValue(dataRow.Col4);

                    rowCur.Cells[7].CellStyle = numberStyle; ;
                    rowCur.Cells[7].SetCellValue(dataRow.Col5);

                    rowCur.Cells[8].CellStyle = numberStyle; ;
                    rowCur.Cells[8].SetCellValue(dataRow.Col6);

                    rowCur.Cells[9].CellStyle = textStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col7);
                    rowCur.Cells[10].SetCellValue(dataRow.Col8);

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col9 == 0 ? 0 : Convert.ToDouble(dataRow.Col9));

                    rowCur.Cells[12].CellStyle = textStyle; ;
                    rowCur.Cells[12].SetCellValue(dataRow.Col10);

                    rowCur.Cells[13].CellStyle = numberStyle; ;
                    rowCur.Cells[13].SetCellValue(dataRow.Col11 == 0 ? 0 : Convert.ToDouble(dataRow.Col11));

                    rowCur.Cells[14].CellStyle = textStyle; ;
                    rowCur.Cells[14].SetCellValue(dataRow.Col12);

                    rowCur.Cells[15].CellStyle = textStyle; ;
                    rowCur.Cells[15].SetCellValue(dataRow.Col13);

                    rowCur.Cells[16].CellStyle = numberStyle; ;
                    rowCur.Cells[16].SetCellValue(dataRow.Col14);

                    rowCur.Cells[17].CellStyle = numberStyle; ;
                    rowCur.Cells[17].SetCellValue(dataRow.Col15);

                    rowCur.Cells[18].CellStyle = numberStyle; ;
                    rowCur.Cells[18].SetCellValue(dataRow.Col16);

                    rowCur.Cells[19].SetCellValue(dataRow.Col17);
                    rowCur.Cells[20].SetCellValue(dataRow.Col18);
                }

                #endregion

                #region Export VK11-FOB (5s)

                var startRowVK11FOB = 3;
                ISheet sheetVK11FOB = workbook.GetSheetAt(13);

                for (var i = 0; i < data.VK11FOB.Count(); i++)
                {
                    var dataRow = data.VK11FOB[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetVK11FOB, startRowVK11FOB++, 17);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);

                    rowCur.Cells[1].SetCellValue(dataRow.ColC == null ? dataRow.ColB : dataRow.ColC);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].SetCellValue(dataRow.Col1);
                    rowCur.Cells[2].CellStyle = textStyle;

                    rowCur.Cells[3].CellStyle = numberStyle;
                    rowCur.Cells[3].SetCellValue(dataRow.Col4);

                    rowCur.Cells[4].CellStyle = numberStyle;
                    rowCur.Cells[4].SetCellValue(dataRow.Col5);

                    rowCur.Cells[5].CellStyle = numberStyle;
                    rowCur.Cells[5].SetCellValue(dataRow.Col6);

                    rowCur.Cells[6].CellStyle = textStyle;
                    rowCur.Cells[6].SetCellValue(dataRow.Col7);

                    rowCur.Cells[7].CellStyle = numberStyle; ;
                    rowCur.Cells[7].SetCellValue(" ");

                    rowCur.Cells[8].CellStyle = numberStyle;
                    rowCur.Cells[8].SetCellValue(dataRow.Col9 == 0 ? 0 : Convert.ToDouble(dataRow.Col9));

                    rowCur.Cells[9].CellStyle = textStyle;
                    rowCur.Cells[9].SetCellValue(dataRow.Col10);

                    rowCur.Cells[10].CellStyle = numberStyle; ;
                    rowCur.Cells[10].SetCellValue(dataRow.Col11 == 0 ? 0 : Convert.ToDouble(dataRow.Col11));

                    rowCur.Cells[11].CellStyle = textStyle;
                    rowCur.Cells[11].SetCellValue(dataRow.Col12);

                    rowCur.Cells[12].CellStyle = textStyle;
                    rowCur.Cells[12].SetCellValue(dataRow.Col13);

                    rowCur.Cells[13].CellStyle = textStyle;
                    rowCur.Cells[13].SetCellValue(dataRow.Col14);

                    rowCur.Cells[14].CellStyle = textStyle;
                    rowCur.Cells[14].SetCellValue(dataRow.Col15);

                    rowCur.Cells[15].CellStyle = numberStyle; ;
                    rowCur.Cells[15].SetCellValue(dataRow.Col16);
                }

                #endregion

                #region Export VK11-TNPP

                var startRowVK11TNPP = 3;
                ISheet sheetVK11TNPP = workbook.GetSheetAt(14);

                for (var i = 0; i < data.VK11TNPP.Count(); i++)
                {
                    var dataRow = data.VK11TNPP[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetVK11TNPP, startRowVK11TNPP++, 20);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);

                    //rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].SetCellValue(dataRow.ColC == null ? dataRow.ColB : dataRow.ColC);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].SetCellValue(dataRow.Col1);
                    rowCur.Cells[2].CellStyle = textStyle;

                    rowCur.Cells[3].CellStyle = numberStyle; ;
                    rowCur.Cells[3].SetCellValue(dataRow.Col2 == 0 ? 0 : Convert.ToDouble(dataRow.Col2));

                    rowCur.Cells[4].CellStyle = numberStyle; ;
                    rowCur.Cells[4].SetCellValue(dataRow.Col3 == 0 ? 0 : Convert.ToDouble(dataRow.Col3));

                    rowCur.Cells[5].CellStyle = numberStyle; ;
                    rowCur.Cells[5].SetCellValue(dataRow.Col4);

                    rowCur.Cells[6].CellStyle = numberStyle; ;
                    rowCur.Cells[6].SetCellValue(dataRow.Col5);

                    rowCur.Cells[7].CellStyle = numberStyle; ;
                    rowCur.Cells[7].SetCellValue(dataRow.Col6);

                    rowCur.Cells[8].CellStyle = textStyle;
                    rowCur.Cells[8].SetCellValue(dataRow.Col7);

                    rowCur.Cells[9].CellStyle = numberStyle;
                    rowCur.Cells[9].SetCellValue(dataRow.Col8);

                    rowCur.Cells[10].CellStyle = numberStyle;
                    rowCur.Cells[10].SetCellValue(dataRow.Col9 == 0 ? 0 : Convert.ToDouble(dataRow.Col9));

                    rowCur.Cells[11].CellStyle = textStyle;
                    rowCur.Cells[11].SetCellValue(dataRow.Col10);

                    rowCur.Cells[12].CellStyle = numberStyle;
                    rowCur.Cells[12].SetCellValue(dataRow.Col11 == 0 ? 0 : Convert.ToDouble(dataRow.Col11));

                    rowCur.Cells[13].CellStyle = textStyle;
                    rowCur.Cells[13].SetCellValue(dataRow.Col12);

                    rowCur.Cells[14].CellStyle = textStyle;
                    rowCur.Cells[14].SetCellValue(dataRow.Col13);

                    rowCur.Cells[15].CellStyle = numberStyle;
                    rowCur.Cells[15].SetCellValue(dataRow.Col14);

                    rowCur.Cells[16].CellStyle = numberStyle;
                    rowCur.Cells[16].SetCellValue(dataRow.Col15);

                    rowCur.Cells[17].CellStyle = numberStyle;
                    rowCur.Cells[17].SetCellValue(dataRow.Col16);

                    rowCur.Cells[18].SetCellValue(dataRow.Col17);
                    rowCur.Cells[19].SetCellValue(dataRow.Col18);

                }

                #endregion

                #region PTS
                var startRowPTS = 4;
                ISheet sheetPTS = workbook.GetSheetAt(16);


                for (var i = 0; i < data.PTS.Count(); i++)
                {
                    var dataRow = data.PTS[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetPTS, startRowPTS++, 20);
                    rowCur.Cells[0].CellStyle = numberStyle; ;
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);

                    rowCur.Cells[1].CellStyle = numberStyle; ;
                    rowCur.Cells[1].SetCellValue(dataRow.Col1);

                    rowCur.Cells[2].CellStyle = numberStyle; ;
                    rowCur.Cells[2].SetCellValue(dataRow.Col2);

                    rowCur.Cells[3].CellStyle = numberStyle; ;
                    rowCur.Cells[3].SetCellValue(dataRow.Col3);

                    rowCur.Cells[4].CellStyle = textStyle; ;
                    rowCur.Cells[4].SetCellValue(dataRow.Col4);

                    rowCur.Cells[5].CellStyle = textStyle; ;
                    rowCur.Cells[5].SetCellValue(dataRow.Col5);

                    rowCur.Cells[6].CellStyle = textStyle; ;
                    rowCur.Cells[6].SetCellValue("");

                    rowCur.Cells[7].CellStyle = textStyle; ;
                    rowCur.Cells[7].SetCellValue("");

                    rowCur.Cells[8].CellStyle = textStyle; ;
                    rowCur.Cells[8].SetCellValue("");

                    rowCur.Cells[9].CellStyle = numberStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col6 == 0 ? 0 : Convert.ToDouble(dataRow.Col6));

                    rowCur.Cells[10].CellStyle = textStyle; ;
                    rowCur.Cells[10].SetCellValue("");

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col7 == 0 ? 0 : Convert.ToDouble(dataRow.Col7));

                    rowCur.Cells[12].CellStyle = textStyle; ;
                    rowCur.Cells[12].SetCellValue(dataRow.Col8);

                    rowCur.Cells[13].CellStyle = textStyle; ;
                    rowCur.Cells[13].SetCellValue("");

                    rowCur.Cells[14].CellStyle = textStyle; ;
                    rowCur.Cells[14].SetCellValue("");

                    rowCur.Cells[15].CellStyle = numberStyle; ;
                    rowCur.Cells[15].SetCellValue(dataRow.Col9);

                    rowCur.Cells[16].CellStyle = numberStyle; ;
                    rowCur.Cells[16].SetCellValue(dataRow.Col10);
                }
                #endregion

                #region Export VK11-BB

                var startRowVK11BB = 3;
                ISheet sheetVK11BB = workbook.GetSheetAt(15);

                for (var i = 0; i < data.VK11BB.Count(); i++)
                {
                    var dataRow = data.VK11BB[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetVK11BB, startRowVK11BB++, 19);
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);

                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].SetCellValue(dataRow.ColC);
                    rowCur.Cells[2].CellStyle = textStyle;

                    rowCur.Cells[3].SetCellValue(dataRow.ColD);
                    rowCur.Cells[3].CellStyle = textStyle;

                    rowCur.Cells[4].CellStyle = numberStyle;
                    rowCur.Cells[4].SetCellValue(dataRow.Col1);

                    rowCur.Cells[5].CellStyle = numberStyle;
                    rowCur.Cells[5].SetCellValue(dataRow.Col2);

                    rowCur.Cells[6].CellStyle = numberStyle;
                    rowCur.Cells[6].SetCellValue(dataRow.Col3);

                    rowCur.Cells[7].CellStyle = textStyle;
                    rowCur.Cells[7].SetCellValue(dataRow.Col4);

                    rowCur.Cells[8].CellStyle = textStyle;
                    rowCur.Cells[8].SetCellValue(dataRow.Col5);

                    rowCur.Cells[9].CellStyle = numberStyle; ;
                    rowCur.Cells[9].SetCellValue(dataRow.Col6 == 0 ? 0 : Convert.ToDouble(dataRow.Col6));

                    rowCur.Cells[10].CellStyle = textStyle;
                    rowCur.Cells[10].SetCellValue(dataRow.Col7);

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col8);

                    rowCur.Cells[12].CellStyle = textStyle;
                    rowCur.Cells[12].SetCellValue(dataRow.Col9);

                    rowCur.Cells[13].CellStyle = textStyle;
                    rowCur.Cells[13].SetCellValue(dataRow.Col10);

                    rowCur.Cells[14].CellStyle = numberStyle; ;
                    rowCur.Cells[14].SetCellValue(dataRow.Col11);

                    rowCur.Cells[15].CellStyle = numberStyle; ;
                    rowCur.Cells[15].SetCellValue(dataRow.Col12);

                    rowCur.Cells[16].CellStyle = numberStyle; ;
                    rowCur.Cells[16].SetCellValue(dataRow.Col13);
                    rowCur.Cells[17].SetCellValue(dataRow.Col14);
                    rowCur.Cells[18].SetCellValue(dataRow.Col15);
                }

                #endregion

                #region tổng hợp

                ISheet sheetTH = workbook.GetSheetAt(17);
                var startRowTH = 4;

                for (var i = 0; i < data.Summary.Count; i++)
                {
                    var dataRow = data.Summary[i];
                    var textStyle = dataRow.IsBold ? styles.TextBold : styles.Text;
                    var numberStyle = dataRow.IsBold ? styles.NumberBold : styles.Number;
                    IRow rowCur = ReportUtilities.CreateRow(ref sheetTH, startRowTH++, 21);

                    rowCur.Cells[0].CellStyle = numberStyle;
                    rowCur.Cells[0].SetCellValue(dataRow.ColA);


                    rowCur.Cells[1].SetCellValue(dataRow.ColB);
                    rowCur.Cells[1].CellStyle = textStyle;

                    rowCur.Cells[2].SetCellValue(dataRow.ColC);
                    rowCur.Cells[2].CellStyle = textStyle;

                    rowCur.Cells[3].CellStyle = textStyle;
                    rowCur.Cells[3].SetCellValue(dataRow.ColD);

                    rowCur.Cells[4].CellStyle = textStyle;

                    rowCur.Cells[5].CellStyle = textStyle;

                    rowCur.Cells[6].CellStyle = numberStyle;
                    rowCur.Cells[6].SetCellValue(dataRow.Col1);

                    rowCur.Cells[7].CellStyle = numberStyle;
                    rowCur.Cells[7].SetCellValue(dataRow.Col2);

                    rowCur.Cells[8].CellStyle = numberStyle;
                    rowCur.Cells[8].SetCellValue(dataRow.Col3);

                    rowCur.Cells[9].CellStyle = textStyle;
                    rowCur.Cells[9].SetCellValue(dataRow.Col4);

                    rowCur.Cells[10].CellStyle = numberStyle;
                    rowCur.Cells[10].SetCellValue(dataRow.Col5);

                    rowCur.Cells[11].CellStyle = numberStyle; ;
                    rowCur.Cells[11].SetCellValue(dataRow.Col6 == 0 ? 0 : Convert.ToDouble(dataRow.Col6));

                    rowCur.Cells[12].CellStyle = textStyle;
                    rowCur.Cells[12].SetCellValue(dataRow.Col7 ?? "VND");

                    rowCur.Cells[13].CellStyle = numberStyle;
                    rowCur.Cells[13].SetCellValue(dataRow.Col8);
                    rowCur.Cells[14].CellStyle = numberStyle;
                    rowCur.Cells[14].SetCellValue(dataRow.Col9);

                    rowCur.Cells[15].CellStyle = textStyle;
                    rowCur.Cells[15].SetCellValue(dataRow.Col10);

                    rowCur.Cells[16].CellStyle = numberStyle;
                    rowCur.Cells[16].SetCellValue(dataRow.Col11);

                    rowCur.Cells[17].CellStyle = numberStyle;
                    rowCur.Cells[17].SetCellValue(dataRow.Col12);

                    rowCur.Cells[18].CellStyle = numberStyle;
                    rowCur.Cells[18].SetCellValue(dataRow.Col13);

                    rowCur.Cells[19].CellStyle = numberStyle;
                    rowCur.Cells[19].SetCellValue(dataRow.Col14);

                    rowCur.Cells[20].CellStyle = numberStyle;
                    rowCur.Cells[20].SetCellValue(dataRow.Col15);

                }
                #endregion


                var outputPath = Path.Combine("Upload/", $"{DateTime.Now:ddMMyyyy_HHmmss}_CSTMGG.xlsx");
                using var outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
                workbook.Write(outFile);

                return outputPath;
            }
            catch
            {
                return string.Empty;
            }
        }

        private void SetCellValues(IRow row, ICellStyle style, dynamic data, int[] indices, decimal[] values,bool isTitle = false)
        {
            for (int i = 0; i < indices.Length; i++)
            {
                row.Cells[indices[i]].CellStyle = style;
                if (isTitle)
                {
                    row.Cells[indices[i]].SetCellValue("");
                }
                else
                {
                    row.Cells[indices[i]].SetCellValue(Convert.ToDouble(values[i]));
                }
            }
        }

        public async Task<string> ExportExcelTrinhKy(string headerId)
        {
            return null;
        }
        //public async Task<string> ExportExcelTrinhKy(string headerId)
        //{
        //    try
        //    {

        //        var data = await GetResult(headerId, -1);
        //        var header = await _dbContext.TblBuCalculateResultList.FindAsync(headerId);
        //        var model = await GetDataInput(headerId);
        //        var goods = await _dbContext.TblMdGoods.ToListAsync();
        //        var NguoiKyTen = await _dbContext.TblMdSigner.FirstOrDefaultAsync(x => x.Code == header.SignerCode);
        //        var A5 = $"  (Kèm theo Công văn số:                        /PLXNA ngày {header.FDate.Day:D2}/{header.FDate.Month:D2}/{header.FDate.Year} của Công ty Xăng dầu Nghệ An)";
        //        var A24 = $" + Căn cứ Quyết định số {header.QuyetDinhSo} ngày {header.FDate.Day:D2}/{header.FDate.Month:D2}/{header.FDate.Year} của Tổng giám đốc Tập đoàn Xăng dầu Việt Nam về việc qui định giá bán xăng dầu; ";
        //        var B25 = $"Mức giá bán đăng ký này có hiệu lực thi hành kể từ 15 giờ 00 ngày {header.FDate.Day} tháng {header.FDate.Month} năm {header.FDate.Year}";
        //        // 1. Đường dẫn file gốc
        //        var filePathTemplate = Path.Combine(Directory.GetCurrentDirectory(), "Template", "TempTrinhKy", "KeKhaiGiaChiTiet.xlsx");

        //        // 2. Tạo thư mục lưu file
        //        var folderName = Path.Combine($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}");
        //        if (!Directory.Exists(folderName))
        //        {
        //            Directory.CreateDirectory(folderName);
        //        }

        //        // 3. Tạo tên file mới
        //        var fileName = $"{DateTime.Now:ddMMyyyy_HHmmss}_KeKhaiGiaChiTiet.xlsx";
        //        var fullPath = Path.Combine(folderName, fileName);

        //        // 4. Copy file từ Template sang Upload
        //        File.Copy(filePathTemplate, fullPath, true);

        //        // 5. Mở file để sửa
        //        IWorkbook workbook;

        //        using (var fs = new FileStream(fullPath, FileMode.Open, FileAccess.ReadWrite))
        //        {
        //            workbook = new XSSFWorkbook(fs);
        //            ISheet sheet = workbook.GetSheetAt(0);
        //            IRow rowA5 = sheet.GetRow(4);
        //            ICell cellA5 = rowA5?.GetCell(0);

        //            if (cellA5 != null)
        //            {
        //                cellA5.SetCellValue(A5);
        //            }

        //            IRow rowA24 = sheet.GetRow(23);
        //            ICell cellA24 = rowA24?.GetCell(0);

        //            if (cellA24 != null)
        //            {
        //                cellA24.SetCellValue(A24);
        //            }

        //            IRow rowB25 = sheet.GetRow(24);
        //            ICell cellB25 = rowB25?.GetCell(1);

        //            if (cellB25 != null)
        //            {
        //                cellB25.SetCellValue(B25);
        //            }
        //            int rowIndex = 10; // Bắt đầu từ row 11 (index = 10)
        //            foreach (var item in data.DLG.Dlg_TDGBL)
        //            {
        //                IRow row = sheet.GetRow(rowIndex); // Chỉ lấy row, không cần CreateRow
        //                if (row != null && !item.Code.Trim().Equals("701001", StringComparison.OrdinalIgnoreCase))
        //                {
        //                    // B11 -> colA
        //                    ICell cellB = row.GetCell(1);
        //                    if (cellB != null)
        //                    {
        //                        cellB.SetCellValue(item.ColA);
        //                    }

        //                    // E11 -> col1
        //                    ICell cellE = row.GetCell(4);
        //                    if (cellE != null && item.Col1.HasValue)
        //                    {
        //                        cellE.SetCellValue((double)item.Col1);
        //                    }

        //                    // F11 -> col2
        //                    ICell cellF = row.GetCell(5);
        //                    if (cellF != null && item.Col2.HasValue)
        //                    {
        //                        cellF.SetCellValue((double)item.Col2);
        //                    }

        //                    // G11 -> tangGiam1_2
        //                    ICell cellG = row.GetCell(6);
        //                    if (cellG != null && item.TangGiam1_2.HasValue)
        //                    {
        //                        cellG.SetCellValue((double)item.TangGiam1_2);
        //                    }

        //                    ICell cellH = row.GetCell(7);
        //                    if (cellH != null)
        //                    {
        //                        if (item.Col1.HasValue && item.Col2.HasValue && item.Col1.Value != 0)
        //                        {
        //                            double rateOfIncreaseAndDecrease = (double)((item.Col2.Value - item.Col1.Value) / item.Col1.Value);
        //                            cellH.SetCellValue(rateOfIncreaseAndDecrease);
        //                        }
        //                        else
        //                        {
        //                            cellH.SetCellValue(0);
        //                        }
        //                    }

        //                }

        //                rowIndex++;
        //            }

        //            int rowIndex2 = 15;
        //            foreach (var item in data.DLG.Dlg_TDGBL)
        //            {
        //                IRow row = sheet.GetRow(rowIndex2); // Chỉ lấy row, không cần CreateRow
        //                if (row != null && !item.Code.Trim().Equals("701001", StringComparison.OrdinalIgnoreCase))
        //                {
        //                    // B11 -> colA
        //                    ICell cellB = row.GetCell(1);
        //                    if (cellB != null)
        //                    {
        //                        cellB.SetCellValue(item.ColA);
        //                    }

        //                    // E11 -> col1
        //                    ICell cellE = row.GetCell(4);
        //                    if (cellE != null && item.Col3.HasValue)
        //                    {
        //                        cellE.SetCellValue((double)item.Col3);
        //                    }

        //                    // F11 -> col2
        //                    ICell cellF = row.GetCell(5);
        //                    if (cellF != null && item.Col4.HasValue)
        //                    {
        //                        cellF.SetCellValue((double)item.Col4);
        //                    }

        //                    // G11 -> tangGiam1_2
        //                    ICell cellG = row.GetCell(6);
        //                    if (cellG != null && item.TangGiam3_4.HasValue)
        //                    {
        //                        cellG.SetCellValue((double)item.TangGiam3_4);
        //                    }

        //                    ICell cellH = row.GetCell(7);
        //                    if (cellH != null)
        //                    {
        //                        if (item.Col3.HasValue && item.Col4.HasValue && item.Col3.Value != 0)
        //                        {
        //                            double rateOfIncreaseAndDecrease = (double)((item.Col4.Value - item.Col3.Value) / item.Col3.Value);
        //                            cellH.SetCellValue(rateOfIncreaseAndDecrease);
        //                        }
        //                        else
        //                        {
        //                            cellH.SetCellValue(0);
        //                        }
        //                    }

        //                }

        //                rowIndex2++;
        //            }

        //            ISheet sheetCheck = workbook.GetSheetAt(0);
        //            IRow rowCheck = sheetCheck.GetRow(4);
        //            ICell cellCheck = rowCheck?.GetCell(0);
        //            Console.WriteLine($"Giá trị trước khi lưu file: {cellCheck?.StringCellValue}");

        //            // Ghi file
        //            using (var fsOut = new FileStream(fullPath, FileMode.Create, FileAccess.Write))
        //            {
        //                workbook.Write(fsOut);
        //                Console.WriteLine("Ghi file thành công");
        //            }
        //            workbook.Close();
        //            return $"{folderName}/{fileName}";
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        Console.WriteLine($"Lỗi: {ex.Message}");
        //        throw;
        //    }
        //}

       
        public async Task<string> SaveFileHistory(MemoryStream outFileStream, string headerId)
        {
            byte[] data = outFileStream.ToArray();
            var path = "";
            using (MemoryStream memoryStream = new MemoryStream(data))
            {
                IFormFile file = ConvertMemoryStreamToIFormFile(memoryStream, "example.txt");
                var folderName = Path.Combine($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}");
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                if (file.Length > 0)
                {
                    var fileName = $"{DateTime.Now.Day}{DateTime.Now.Month}{DateTime.Now.Year}_{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}_CoSoTinhMucGiamGia.xlsx";
                    var fullPath = Path.Combine(pathToSave, fileName);
                    using (var stream = new FileStream(fullPath, FileMode.Create))
                    {
                        file.CopyTo(stream);
                    }
                    path = $"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}/{fileName}";
                    _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = headerId,
                        Name = fileName,
                        Type = "xlsx",
                        Path = path
                    });
                    await _dbContext.SaveChangesAsync();
                }
            }
            return path;
        }

        public static IFormFile ConvertMemoryStreamToIFormFile(MemoryStream memoryStream, string fileName)
        {
            memoryStream.Position = 0; // Reset the stream position to the beginning
            IFormFile formFile = new FormFile(memoryStream, 0, memoryStream.Length, "file", fileName)
            {
                Headers = new HeaderDictionary(),
                ContentType = "application/octet-stream"
            };
            return formFile;

        }

        public async Task<string> GenarateWordTrinhKy(string headerId, string nameTemp)
        {
            #region Tạo 1 file word mới từ file template    
            var filePathTemplate = Directory.GetCurrentDirectory() + $"/Template/TempTrinhKy/{nameTemp}.docx";
            var folderName = Path.Combine($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}");
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
            var data = await GetResult(headerId, -1);

            var header = await _dbContext.TblBuCalculateResultList.FindAsync(headerId);
            var model = await GetDataInput(headerId);
            var goods = await _dbContext.TblMdGoods.ToListAsync();
            var NguoiKyTen = await _dbContext.TblMdSigner.FirstOrDefaultAsync(x => x.Code == header.SignerCode);
            var f_date = $"{header.FDate.Day:D2} tháng {header.FDate.Month:D2} năm {header.FDate.Year}";
            var date = header.FDate.ToString("dd/MM/yyyy");
            var f_date_hour = $"kể từ {header.FDate.Hour:D2} giờ {header.FDate.Minute:D2} ngày {header.FDate.Day:D2} tháng {header.FDate.Month:D2} năm {header.FDate.Year}";

            var OldCalculate = await _dbContext.TblBuCalculateResultList
                                    .Where(x => x.FDate < header.FDate)
                                    .Where(x => x.Status == "04")
                                    .OrderByDescending(x => x.FDate)
                                    .FirstOrDefaultAsync();
            var model_Old = await GetDataInput(OldCalculate.Code);
            var data__DLG_DLG_2_Old = new List<DLG_2>();
            var dataVCL = await _dbContext.TblInVinhCuaLo.Where(x => x.HeaderCode == OldCalculate.Code).ToListAsync();
            if (dataVCL.Count() == 0)
            {
                return "lỗi k có dữ liệu dataVCL hoặc dataHSMH";
            }
            foreach (var g in goods)
            {
                var vcl = dataVCL.Where(x => x.GoodsCode == g.Code).ToList();

                data__DLG_DLG_2_Old.Add(new DLG_2
                {
                    Code = g.Code,
                    Col1 = g.Name,
                    Col2 = vcl.Sum(x => x.GblV2),
                });

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
                var sortedHS2 = model.HS2.OrderBy(x => int.Parse(x.GoodsCode)).ToList();
                var dlg1 = data.DLG.Dlg_1;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##F_DATE@@":
                                var text = $"ngày {header.FDate.Day} tháng {header.FDate.Month} năm {header.FDate.Year}";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, text);
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
                                    var goodsList = goods.Where(x => x.IsActive == true)
                                                         .OrderByDescending(x => x.ThueBvmt)
                                                         .ToList();
                                    foreach (var i in goodsList)
                                    {
                                        var HS2Item = model.HS2.FirstOrDefault(x => x.GoodsCode == i.Code);
                                        if (HS2Item != null)
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.Name, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(HS2Item.Gny.ToString("N0"), true, 26, false, "2400"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
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
                                    foreach (var i in data.DLG.Dlg_2)
                                    {
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.Col1, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(i.Col2.ToString("N0"), true, 26, false, "2400"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                            table.Append(row);
                                            o++;
                                        }
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
                var dlg5 = data.DLG.Dlg_5;
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

                                    rowHeader.Append(cell1);
                                    rowHeader.Append(cell2);
                                    rowHeader.Append(cell3);
                                    rowHeader.Append(cell4);

                                    table.Append(rowHeader);
                                    #endregion

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in dlg5)
                                    {
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.ColA, true, 26, true, "1"));
                                            row.Append(CreateCell(i.ColB, true, 26, true));
                                            row.Append(CreateCell(i.Col5.ToString("N0"), true, 26));
                                            row.Append(CreateCell("đ/ lít thực tế", true, 26));
                                            table.Append(row);
                                            o++;
                                        }
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
                var dlg6 = data.DLG.Dlg_6;
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

                                    rowHeader.Append(cell1);
                                    rowHeader.Append(cell2);
                                    rowHeader.Append(cell3);
                                    rowHeader.Append(cell4);

                                    table.Append(rowHeader);
                                    #endregion

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in dlg6)
                                    {
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.ColA, true, 26, true, "1"));
                                            row.Append(CreateCell(i.ColB, true, 26, true, "3000"));
                                            row.Append(CreateCell(i.Col6.ToString("N0"), true, 26, true, "3000"));
                                            row.Append(CreateCell("đ/ lít thực tế", true, 26, true, "3000"));
                                            table.Append(row);
                                            o++;
                                        }
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
                var dlg7 = data.DLG.Dlg_7;
                var dlg8 = data.DLG.Dlg_8;
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

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
                            case "##F_DATE_HOURE@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, f_date_hour);
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
                                    foreach (var i in dlg7.Where(x => x.Type == "TT"))
                                    {
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.ColA, false, 26, false, "3000")); // Tên mặt hàng
                                            row.Append(CreateCell(i.Col1.ToString("N0"), false, 26, false, "2082")); // LG cũ
                                            row.Append(CreateCell(i.Col2.ToString("N0"), false, 26, false, "2082")); // LG mới
                                            row.Append(CreateCell(i.TangGiam1_2.ToString("N0"), false, 26, false, "2082"));
                                            table.Append(row);
                                        }

                                    }

                                    // Thêm dòng tiêu đề "Các vùng thị trường còn lại"
                                    TableRow regionRow2 = new TableRow();
                                    regionRow2.Append(CreateCell("Các vùng thị trường còn lại", true, 26, true, "2082", 4, 1)); // Gộp 4 cột
                                    table.Append(regionRow2);

                                    // Duyệt danh sách dlg7, in từng mặt hàng thuộc vùng còn lại
                                    foreach (var i in dlg7.Where(x => x.Type == "OTHER"))
                                    {
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell(i.ColA, false, 26, false, "3000")); // Tên mặt hàng
                                            row.Append(CreateCell(i.Col1.ToString("N0"), false, 26, false, "2082")); // LG cũ
                                            row.Append(CreateCell(i.Col2.ToString("N0"), false, 26, false, "2082")); // LG mới
                                            row.Append(CreateCell(i.TangGiam1_2.ToString("N0"), false, 26, false, "2082"));
                                            table.Append(row);
                                        }

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
                                    var o = 1;
                                    foreach (var i in dlg8.Where(x => x.Type == "TT"))
                                    {
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();

                                            row.Append(CreateCell(i.ColA, false, 26, false, "3000")); // Tên mặt hàng
                                            row.Append(CreateCell(i.Col1.ToString("N0"), false, 26, false, "2082")); // LG cũ
                                            row.Append(CreateCell(i.Col2.ToString("N0"), false, 26, false, "2082")); // LG mới
                                            row.Append(CreateCell(i.TangGiam1_2.ToString("N0"), false, 26, false, "2082"));
                                            table.Append(row);
                                        }

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
                var sortedHS2 = model.HS2.OrderBy(x => int.Parse(x.GoodsCode)).ToList();
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

                                    Console.WriteLine("Dữ liệu model đợt cũ" + model_Old.Header.Name);
                                    #region Gendata table
                                    var o = 1;
                                    var goodsList = goods.Where(x => x.IsActive == true)
                                                         .OrderByDescending(x => x.ThueBvmt)
                                                         .ToList();
                                    foreach (var i in goodsList)
                                    {

                                        var HS2Item = model.HS2.FirstOrDefault(x => x.GoodsCode == i.Code);
                                        var HS2Item_Old = model_Old.HS2.FirstOrDefault(x => x.GoodsCode == i.Code);
                                        Console.WriteLine("Dữ liệu model đợt cũ" + HS2Item_Old.Gny);
                                        if (HS2Item != null && HS2Item_Old != null)
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.Name, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(HS2Item.Gny.ToString("N0"), true, 26, false, "2400"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                            row.Append(CreateCell(HS2Item.Gny != HS2Item_Old.Gny ? "Thay đổi" : "Không thay đổi", false, 26, false, "2400"));
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

                                    for (int index = 0; index < data.DLG.Dlg_2.Count; index++)
                                    {
                                        var i = data.DLG.Dlg_2[index];
                                        var i_Old = data__DLG_DLG_2_Old.FirstOrDefault(x => x.Col1 == i.Col1);
                                        if (i.Code != "701001")
                                        {
                                            TableRow row = new TableRow();
                                            row.Append(CreateCell("+ " + i.Col1, true, 26, false, "3500"));
                                            row.Append(CreateCell(":", false, 26, true, "1"));
                                            row.Append(CreateCell(i.Col2.ToString("N0"), true, 26, false, "2400"));
                                            row.Append(CreateCell("đ/lít thực tế", false, 26, false, "2400"));
                                            row.Append(CreateCell(i.Col2 != i_Old.Col2 ? "Thay đổi" : "Không thay đổi", false, 26, false, "2400"));
                                            table.Append(row);
                                            o++;
                                        }
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
            #endregion






            #endregion

            return $"{folderName}/{fileName}";
        }

        public async Task<string> GenarateWord(List<string> lstCustomerChecked, string headerId)
        {
            #region Tạo 1 file word mới từ file template
            var filePathTemplate = Directory.GetCurrentDirectory() + "/Template/ThongBaoGia.docx";
            var folderName = Path.Combine($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}");
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
            var data = await GetResult(headerId, -1);
            var header = await _dbContext.TblBuCalculateResultList.FindAsync(headerId);
            foreach (var code in lstCustomerChecked)
            {
                var d = data.VK11BB.Where(x => x.Col2 == code).ToList();
                var c = await _dbContext.TblMdCustomer.FindAsync(code);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(fullPath, true))
                {
                    MainDocumentPart mainPart = doc.MainDocumentPart;
                    DocumentFormat.OpenXml.Wordprocessing.Body body = mainPart.Document.Body;

                    foreach (var t in lstTextElement)
                    {
                        switch (t)
                        {
                            case "##DATE@@":
                                var text = $"{header.FDate.Hour}h{header.FDate.Minute} ngày {header.FDate.Day} tháng {header.FDate.Month} năm {header.FDate.Year}";
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, text);
                                break;
                            case "##COMPANY@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, c.Name);
                                break;
                            case "##ADDRESS@@":
                                wordDocumentService.ReplaceStringInWordDocumennt(doc, t, c.Address);
                                break;
                            case "##TABLE@@":
                                Paragraph paragraph = body.Descendants<Paragraph>()
                                           .FirstOrDefault(p => p.InnerText.Contains("##TABLE@@"));
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
                                    TableRow rowHeader = new TableRow();


                                    TableCell CreateHeaderCell(string text, int gridSpan, int fontSize = 16)
                                    {
                                        TableCell cell = new TableCell();
                                        Paragraph paragraph = new Paragraph(new Run(new Text(text)));
                                        paragraph.ParagraphProperties = new ParagraphProperties(new Justification() { Val = JustificationValues.Center });
                                        Run run = paragraph.Elements<Run>().First();
                                        run.RunProperties = new RunProperties(
                                           new Bold(),
                                           new FontSize() { Val = new StringValue(fontSize.ToString()) }
                                       );

                                        cell.Append(paragraph);
                                        TableCellProperties cellProperties = new TableCellProperties(
                                            new TableCellVerticalAlignment() { Val = TableVerticalAlignmentValues.Center }
                                        );
                                        if (gridSpan > 1)
                                        {
                                            cellProperties.Append(new GridSpan() { Val = gridSpan });
                                        }
                                        cell.Append(cellProperties);
                                        return cell;
                                    }


                                    TableCell CreateCell(string text, bool isBold, int fontSize = 26)
                                    {
                                        Run run = new Run(new Text(text));
                                        RunProperties runProperties = new RunProperties(
                                            new FontSize() { Val = new StringValue(fontSize.ToString()) }
                                        );
                                        if (isBold)
                                        {
                                            runProperties.AppendChild(new Bold());
                                        }
                                        run.RunProperties = runProperties;
                                        Paragraph paragraph = new Paragraph(run);
                                        return new TableCell(paragraph);
                                    }
                                    TableCell cell1 = CreateHeaderCell("STT", 1, 26);
                                    TableCell cell2 = CreateHeaderCell("Mặt hàng", 1, 26);
                                    TableCell cell3 = CreateHeaderCell("Điểm giao hàng", 1, 26);
                                    TableCell cell4 = CreateHeaderCell("Đơn giá", 2, 26);
                                    rowHeader.Append(cell1);
                                    rowHeader.Append(cell2);
                                    rowHeader.Append(cell3);
                                    rowHeader.Append(cell4);
                                    table.Append(rowHeader);
                                    #endregion

                                    #region Gendata table
                                    var o = 1;
                                    foreach (var i in d)
                                    {
                                        TableRow row = new TableRow();
                                        row.Append(CreateCell(o.ToString(), false, 26));
                                        row.Append(CreateCell(i.ColD, false, 26));
                                        row.Append(CreateCell(i.ColC, false, 26));
                                        row.Append(CreateCell(i.Col6.ToString("N0"), false, 26));
                                        row.Append(CreateCell("Đ/lít tt", false, 26));
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

                if (code != lstCustomerChecked.LastOrDefault())
                {
                    AppendWordFilesToNewDocument(filePathTemplate, fullPath);
                }
            }
            #endregion

            return $"{folderName}/{fileName}";
        }

        public async Task<string> GenarateFile(List<string> lstCustomerChecked, string type, string headerId, CalculateResultModel data)
        {

            if (type == "WORD")
            {
                var path = await GenarateWord(lstCustomerChecked, headerId);
                _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                {
                    Code = Guid.NewGuid().ToString(),
                    HeaderCode = headerId,
                    Name = path.Replace($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}/", ""),
                    Type = "docx",
                    Path = path
                });
                await _dbContext.SaveChangesAsync();
                return path;
            }
            if (type == "WORDTRINHKY")
            {
                foreach (var n in lstCustomerChecked)
                {
                    if (n == "KeKhaiGiaChiTiet")
                    {
                        var path = await ExportExcelTrinhKy(headerId);
                        _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                        {
                            Code = Guid.NewGuid().ToString(),
                            HeaderCode = headerId,
                            Name = path.Replace($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}/", ""),
                            Type = "xlsx",
                            Path = path
                        });
                        await _dbContext.SaveChangesAsync();
                    }
                    else
                    {
                        var path = await GenarateWordTrinhKy(headerId, n);
                        _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                        {
                            Code = Guid.NewGuid().ToString(),
                            HeaderCode = headerId,
                            Name = path.Replace($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}/", ""),
                            Type = "docx",
                            Path = path
                        });
                        await _dbContext.SaveChangesAsync();
                    }
                }
                return null;
            }
            if(type == "EXCEL")
            {
               
                    var path = await ExportExcelPlus(headerId, data);
                    _dbContext.TblBuHistoryDownload.Add(new TblBuHistoryDownload
                    {
                        Code = Guid.NewGuid().ToString(),
                        HeaderCode = headerId,
                        Name = path.Replace($"Upload/", ""),
                        Type = "xlsx",
                        Path = path
                    });
                    await _dbContext.SaveChangesAsync();
                    return path;
               

            }
            else
            {
                var w = await GenarateWord(lstCustomerChecked, headerId);
                var pathWord = Directory.GetCurrentDirectory() + "/" + w;
                Aspose.Words.Document doc = new Aspose.Words.Document(pathWord);
                var folderName = Path.Combine($"Upload/{DateTime.Now.Year}/{DateTime.Now.Month}");
                if (!Directory.Exists(folderName))
                {
                    Directory.CreateDirectory(folderName);
                }
                var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
                var fileName = $"{DateTime.Now.Day}{DateTime.Now.Month}{DateTime.Now.Year}_{DateTime.Now.Hour}{DateTime.Now.Minute}{DateTime.Now.Second}_ThongBaoGia.pdf";
                var fullPath = Path.Combine(pathToSave, fileName);
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
                return $"{folderName}/{fileName}";
            }
        }

        static void AppendWordFilesToNewDocument(string directoryPath, string newWordFilePath)
        {
            using (WordprocessingDocument sourceDocument = WordprocessingDocument.Open(directoryPath, false))
            {
                DocumentFormat.OpenXml.Wordprocessing.Body sourceBody = sourceDocument.MainDocumentPart.Document.Body;

                using (WordprocessingDocument destinationDocument = WordprocessingDocument.Open(newWordFilePath, true))
                {
                    DocumentFormat.OpenXml.Wordprocessing.Body destinationBody = destinationDocument.MainDocumentPart.Document.Body;
                    foreach (var element in sourceBody.Elements())
                    {
                        destinationBody.Append(element.CloneNode(true));
                    }
                    destinationDocument.MainDocumentPart.Document.Save();
                }
            }

        }
        static Table CreateSampleTable()
        {
            Table table = new Table();
            TableProperties tblProp = new TableProperties(
                new TableBorders(
                    new TopBorder { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                    new BottomBorder { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                    new LeftBorder { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                    new RightBorder { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                    new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 },
                    new InsideVerticalBorder { Val = new EnumValue<BorderValues>(BorderValues.Thick), Size = 8 }));
            table.AppendChild(tblProp); TableRow tr = new TableRow();
            TableCell tc1 = new TableCell(new Paragraph(new Run(new Text("Cell 1"))));
            TableCell tc2 = new TableCell(new Paragraph(new Run(new Text("Cell 2"))));
            tr.Append(tc1); tr.Append(tc2);
            table.Append(tr); return table;
        }

    }
}
public static class ExcelNPOIExtentions
{
    public static ICellStyle SetCellStyleText(IWorkbook workbook)
    {
        ICellStyle style = workbook.CreateCellStyle();
        IFont font = workbook.CreateFont();
        font.FontName = "Segoe UI";
        font.FontHeightInPoints = 12;
        style.SetFont(font);
        style.WrapText = true;
        style.Alignment = HorizontalAlignment.Left;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        return style;
    }

    public static ICellStyle SetCellStyleTextBold(IWorkbook workbook)
    {
        ICellStyle style = workbook.CreateCellStyle();
        IFont font = workbook.CreateFont();
        font.FontName = "Segoe UI";
        font.FontHeightInPoints = 12;
        font.IsBold = true;
        style.SetFont(font);
        style.WrapText = true;
        style.Alignment = HorizontalAlignment.Left;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        return style;
    }

    public static ICellStyle SetCellStyleTextSign(IWorkbook workbook,bool isCenter,bool isFill,bool isBold, bool isVerticalAlignment = false)
    {
        ICellStyle style = workbook.CreateCellStyle();
        IFont font = workbook.CreateFont();
        font.FontName = "Segoe UI";
        font.FontHeightInPoints = 12;
        font.IsBold = isBold?true:false;
        style.SetFont(font);
        if (isFill)
        {
        style.FillForegroundColor = IndexedColors.White.Index;
        style.FillPattern = FillPattern.SolidForeground;
        }
        style.Alignment = isCenter ? HorizontalAlignment.Center : HorizontalAlignment.Left;
        style.VerticalAlignment = isVerticalAlignment ? VerticalAlignment.Center: VerticalAlignment.Bottom;
        return style;
    }

    public static ICellStyle SetCellStyleNumber(IWorkbook workbook)
    {
        ICellStyle style = workbook.CreateCellStyle();
        IFont font = workbook.CreateFont();
        font.FontName = "Segoe UI";
        font.FontHeightInPoints = 12;
        style.SetFont(font);
        style.Alignment = HorizontalAlignment.Right;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        style.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");
        return style;
    }

    public static ICellStyle SetCellStyleNumberBold(IWorkbook workbook)
    {
        ICellStyle style = workbook.CreateCellStyle();
        IFont font = workbook.CreateFont();
        font.FontName = "Segoe UI";
        font.FontHeightInPoints = 12;
        font.IsBold = true;
        style.SetFont(font);
        style.Alignment = HorizontalAlignment.Right;
        style.BorderTop = BorderStyle.Thin;
        style.BorderBottom = BorderStyle.Thin;
        style.BorderLeft = BorderStyle.Thin;
        style.BorderRight = BorderStyle.Thin;
        style.DataFormat = workbook.CreateDataFormat().GetFormat("#,##0");
        return style;
    }
}
