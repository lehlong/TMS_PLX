using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using Microsoft.Extensions.Logging;
using Microsoft.EntityFrameworkCore;
using DMS.CORE;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Net;
using SendGrid;
using SendGrid.Helpers.Mail;
using Microsoft.Extensions.Configuration;
using static Org.BouncyCastle.Math.EC.ECCurve;
using DocumentFormat.OpenXml.Office2013.Drawing.ChartStyle;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DMS.BUSINESS.Services.BackgroundHangfire
{
    public class SMSInfo
    {
        public string? UrlSMS { get; set; }
        public string? Username { get; set; }
        public string? Password { get; set; }
        public string? CpCode { get; set; }
        public string? ServiceId { get; set; }
    }
    public class BackgroundJobService
    {
        private readonly AppDbContext _dbContext;
        private SMSInfo _config;
        public BackgroundJobService(AppDbContext dbContext)
        {
            _dbContext = dbContext;
            _config = new SMSInfo
            {
                UrlSMS = "http://ams.tinnhanthuonghieu.vn:8009/bulkapi",
                Username = "smsbrand_xangdauna",
                Password = "xd@258369",
                CpCode = "XANGDAUNA",
                ServiceId = "CtyXdauN.an"
            };
        }

        public async Task SendSMSAsync()
        {
            try
            {
                _dbContext.ChangeTracker.Clear();
                var lstSMS = await _dbContext.TblCmNotifySms.Where(x => x.IsSend == "N").ToListAsync();
                foreach (var s in lstSMS)
                {
                    var status = await SendSMS(ConvertPhoneNumber(s.PhoneNumber), s.Contents);
                    if (status)
                    {
                        s.IsSend = "Y";
                        _dbContext.TblCmNotifySms.Update(s);
                        _dbContext.SaveChanges();
                        Console.WriteLine("Gửi tin nhắn thành công!");
                    }
                    else
                    {
                        Console.WriteLine("Lỗi không gửi được tin nhắn");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }

        public string ConvertPhoneNumber(string phoneNumber)
        {
            if (!string.IsNullOrEmpty(phoneNumber) && phoneNumber.StartsWith("0") && phoneNumber.Length > 1)
            {
                return "84" + phoneNumber.Substring(1);
            }
            return phoneNumber;
        }

        public async Task<bool> SendSMS(string phone, string content)
        {
            try
            {
                string soapRequest = $@"
                    <soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:impl='http://impl.bulkSms.ws/'>
                       <soapenv:Header/>
                       <soapenv:Body>
                          <impl:wsCpMt>
                             <User>{_config.Username}</User>
                             <Password>{_config.Password}</Password>
                             <CPCode>{_config.CpCode}</CPCode>
                             <RequestID>1</RequestID>
                             <UserID>{phone}</UserID>
                             <ReceiverID>{phone}</ReceiverID>
                             <ServiceID>{_config.ServiceId}</ServiceID>
                             <CommandCode>bulksms</CommandCode>
                             <Content>{content}</Content>
                             <ContentType>1</ContentType>
                          </impl:wsCpMt>
                       </soapenv:Body>
                    </soapenv:Envelope>";
                using (HttpClient client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Add("SOAPAction", "wsCpMt");
                    HttpContent contentData = new StringContent(soapRequest, Encoding.UTF8, "text/xml");

                    HttpResponseMessage response = await client.PostAsync(_config.UrlSMS, contentData);
                    var res = await response.Content.ReadAsStringAsync();
                    if (!string.IsNullOrEmpty(res))
                    {
                        if (res.Contains("<result>1</result>"))
                        {
                            return true;
                        }
                    }
                }
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Lỗi không gửi được SMS! Chi tiết {ex.Message}");
                return false;
            }

        }
    }
}
