using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Services.BU;
using DMS.BUSINESS.Services.Report;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using PROJECT.Service.Extention;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http.Headers;
using DMS.CORE;
using DMS.BUSINESS.Dtos.MD;
using DMS.CORE.Entities.MD;

namespace DMS.API.Controllers.Report
{
    [Route("api/[controller]")]
    [ApiController]
    public class ReportController(IReportService service) : ControllerBase
    {
        public readonly IReportService _service = service;

        public AppDbContext _context;

        [HttpGet("GetListTemplate/{yearValue}/{auditValue}")]
        public async Task<IActionResult> Search(string yearValue, string auditValue)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetListTemplate(yearValue, auditValue);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0001", _service);
            }
            return Ok(transferObject);
        }
        [HttpGet("GetTemplate/{id}/{year}/{audit}")]
        public async Task<IActionResult> GetTemplate(string id, string year, string audit)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetTemplate(id, year, audit);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0001", _service);
            }
            return Ok(transferObject);
        }

        [HttpGet("GetListElement/{fileId}")]
        public async Task<IActionResult> GetListElement(string fileId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetListElement(fileId);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0001", _service);
            }
            return Ok(transferObject);
        }

        [HttpPost("Upload")]
        public async Task<IActionResult> Insert(IFormFile file, string? moduleType, Guid? referenceId, string auditValue, string yearValue)
        {
            var folderName = Path.Combine($"wwwroot");
            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            if (file.Length > 0)
            {
                var fileName = Guid.NewGuid().ToString() + ".docx";
                var fullPath = Path.Combine(pathToSave, fileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                await _service.ProcessFile(fullPath, auditValue, yearValue);
                var url = $"{fileName}";

                var transferObject = new TransferObject();
                if (_service.Status)
                {
                    transferObject.Data = url;
                }
                else
                {
                    transferObject.Status = false;
                    transferObject.MessageObject.MessageType = MessageType.Error;
                    transferObject.GetMessage("2000", _service);
                }
                return Ok(transferObject);
            }
            else
            {
                return BadRequest();
            }
        }

        [HttpPost("UploadTemplate")]
        public async Task<IActionResult> UploadTemplate(IFormFile file, string? moduleType, Guid? referenceId, string auditValue, string yearValue)
        {
            var folderName = Path.Combine($"wwwroot");
            if (!Directory.Exists(folderName))
            {
                Directory.CreateDirectory(folderName);
            }
            var pathToSave = Path.Combine(Directory.GetCurrentDirectory(), folderName);
            if (file.Length > 0)
            {
                var fileName = Guid.NewGuid().ToString() + ".docx";
                var fileOldName = file.FileName;
                var fileExt = file.GetType().Name;
                var fileSize = file.Length.ToString();
                var netWorkPath = pathToSave;
                var fullPath = Path.Combine(pathToSave, fileName);
                using (var stream = new FileStream(fullPath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }
                await _service.SaveReportTemplate(fullPath, auditValue, yearValue, fileOldName, fileExt, fileSize, fileName, netWorkPath);

                var transferObject = new TransferObject();
                if (_service.Status)
                {
                    transferObject.Status = true;
                    transferObject.MessageObject.MessageType = MessageType.Success;
                    transferObject.GetMessage("0100", _service);
                }
                else
                {
                    transferObject.Status = false;
                    transferObject.MessageObject.MessageType = MessageType.Error;
                    transferObject.GetMessage("2000", _service);
                }
                return Ok(transferObject);
            }
            else
            {
                return BadRequest();
            }
        }
        [HttpGet("GetListOrg")]
        public async Task<IActionResult> GetListOrg([FromQuery] string fileId, [FromQuery] string textElement)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetListOrg(fileId, textElement);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }
        [HttpGet("GetListOpinion")]
        public async Task<IActionResult> GetListOpinion([FromQuery] string fileId, [FromQuery] string textElement, [FromQuery] string OrgCode, [FromQuery] string timeYear, [FromQuery] string auditPeriod)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetListOpinion(fileId, textElement,OrgCode, timeYear , auditPeriod);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }
        [HttpPost("SaveTemplateReport")]
        public async Task<IActionResult> SaveTemplateReport(TemplateReportMappingDto dto)
        {
            var transferObject = new TransferObject();
            var result = _service.SaveTemplateReport(dto);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0001", _service);
            }
            return Ok(transferObject);

        }
    }
}
