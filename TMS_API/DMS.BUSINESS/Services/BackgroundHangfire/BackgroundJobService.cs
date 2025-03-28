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

namespace DMS.BUSINESS.Services.BackgroundHangfire
{
    public class BackgroundJobService
    {
        private readonly AppDbContext _dbContext;

        public BackgroundJobService(AppDbContext dbContext)
        {
            _dbContext = dbContext;
        }

        public async Task SendEmailsAsync(string apikeySendGrid)
        {
            try
            {


                var dataTable = await _dbContext.TblCmNotifiEmail.Where(x => x.IsSend == "N").ToListAsync();
                var apiKey = Environment.GetEnvironmentVariable(apikeySendGrid);
                var client = new SendGridClient(apiKey);
                var from = new EmailAddress("clonegma1l11@gmail.com", "MailUser");

                foreach (var item in dataTable)
                {
                    try
                    {

                       
                        var to = new EmailAddress(item.Email, "Example User");
                        var plainTextContent = "đây là mail được gửi bởi D2s";
                       
                        var msg = MailHelper.CreateSingleEmail(from, to, item.Subject, plainTextContent, item.Contents);
                        var response = await client.SendEmailAsync(msg);

                        var entity = _dbContext.TblCmNotifiEmail.Find(item.Id);
                        entity.IsSend = "Y";
                        _dbContext.TblCmNotifiEmail.Update(entity);
                        await _dbContext.SaveChangesAsync();

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex);
                        var entity = _dbContext.TblCmNotifiEmail.Find(item.Id);
                        entity.NumberRetry = item.NumberRetry +1;
                        _dbContext.TblCmNotifiEmail.Update(entity);
                        await _dbContext.SaveChangesAsync();

                        continue;
                    }
                }


                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }
        }
    }
}
