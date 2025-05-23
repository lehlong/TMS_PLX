using System.Linq;
using System.Text;
using Microsoft.EntityFrameworkCore;
using DMS.CORE;
using System.Data;
using System.Net.Mail;
using System.Net;
using DocumentFormat.OpenXml.Office.CustomUI;
using Microsoft.IdentityModel.Tokens;

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
    public class EmmailInfo
    {
        public int? port { get; set; }
        public string ? host { get; set; }
        public string? pass { get; set; }
        public string? Email { get; set; }
    }
    public class BackgroundJobService
    {
        private readonly AppDbContext _dbContext;
        private SMSInfo _config;
        private EmmailInfo _congifEmail;
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
            _congifEmail = new EmmailInfo
            {
                port = 20,
                Email= "longlh@petrolimex.com.vn",
                host= "mail.petrolimex.com.vn"
            };

           
        }
                //UrlSMS = "http://ams.tinnhanthuonghieu.vn:8009/bulkapi",

        public async Task SendSMSAsync()
        {
            try
            {
                _dbContext.ChangeTracker.Clear();
                var lstSMS = await _dbContext.TblCmNotifySms.Where(x => x.IsSend == "N" && x.NumberRetry < 3 && x.IsSend != "K").ToListAsync();
                foreach (var s in lstSMS)
                {
                    var status = await SendSMS(ConvertPhoneNumber(s.PhoneNumber), s.Contents);
                    if (status)
                    {
                        s.NumberRetry = s.NumberRetry + 1;
                        s.IsSend = "Y";
                        _dbContext.TblCmNotifySms.Update(s);
                        _dbContext.SaveChanges();
                        Console.WriteLine("Gửi tin nhắn thành công!");
                    }
                    else
                    {
                        s.NumberRetry = s.NumberRetry + 1;
                        s.IsSend = "N";
                        _dbContext.TblCmNotifySms.Update(s);
                        _dbContext.SaveChanges();
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

        public async Task<bool> SendMail(string toEmail, string subject, string body,string path)
        {
            try
            {

                var smtpClient = new SmtpClient(_congifEmail.host)
                {
                    Port = 587,
                    Credentials = new NetworkCredential(_congifEmail.Email, _congifEmail.pass),
                    EnableSsl = true,
                };
                var mailMessage = new MailMessage
                {
                    From = new MailAddress(_congifEmail.Email, "TMS"),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = true,
                };
                if (!string.IsNullOrEmpty(path) && File.Exists(path))
                {
                    var attachment = new Attachment(path);
                    mailMessage.Attachments.Add(attachment);
                }

                mailMessage.To.Add("Somot1pro@gmail.com");
                await smtpClient.SendMailAsync(mailMessage);
                return true;

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to send email. Error: {ex.Message}");
                return false;
            }
        }

        public async Task SendMailAsync()
        {
            try
            {
                _dbContext.ChangeTracker.Clear();
                var lstEmail = await _dbContext.TblCmNotifiEmail.Where(x => x.IsSend == "N" && x.NumberRetry<3).ToListAsync();
                var lstFiledowload = await _dbContext.TblBuHistoryDownload
     .Where(x => x.CustomerCode != null && x.CustomerCode != "")
     .ToListAsync();
                foreach (var s in lstEmail)
                {
                   var Path= lstFiledowload.Where(x => x.CustomerCode == s.CustomerCode&&x.HeaderCode==s.HeaderId).OrderBy(x=>x.CreateDate).Select(x => x.Path).FirstOrDefault();
                    var status = await SendMail(s.Email, s.Subject, s.Contents, Path);
                    if (status)
                    {
                        s.IsSend = "Y";
                       
                        s.NumberRetry = s.NumberRetry + 1;
                        _dbContext.TblCmNotifiEmail.Update(s);
                       
                        Console.WriteLine("Gửi email thành công!");
                    }
                    else
                    {
                      
                        s.NumberRetry = s.NumberRetry + 1;
                        _dbContext.TblCmNotifiEmail.Update(s);
                        Console.WriteLine("Lỗi không gửi được email");
                    }
                    _dbContext.SaveChanges();

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
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
