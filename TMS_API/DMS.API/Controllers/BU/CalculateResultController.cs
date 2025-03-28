using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Models;
using DMS.BUSINESS.Services.BU;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using NPOI.HSSF.Record.Chart;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    public class CalculateResultController : ControllerBase
    {
        public readonly ICalculateResultService _service;
        public CalculateResultController(ICalculateResultService service)
        {
            _service = service;
        }

        [HttpGet("GetCalculateResult")]
        //[Authorize]
        public async Task<IActionResult> GetCalculateResult([FromQuery] string code, [FromQuery] int tab)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetResult(code, tab);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                //transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }

        [HttpGet("GetDataInput")]
       // [Authorize]
        public async Task<IActionResult> GetDataInput([FromQuery] string code)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetDataInput(code);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                //transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }

        [HttpGet("GetHistoryAction")]
        [Authorize]
        public async Task<IActionResult> GetHistoryAction([FromQuery] string code)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetHistoryAction(code);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                //transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }

        [HttpGet("GetHistoryFile")]
        [Authorize]
        public async Task<IActionResult> GetHistoryFile([FromQuery] string code)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetHistoryFile(code);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                //transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }

        [HttpGet("GetCustomer")]
        [Authorize]
        public async Task<IActionResult> GetCustomer()
        {
            var transferObject = new TransferObject();
            var result = await _service.GetCustomer();
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                //transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }

        [HttpPost("UpdateDataInput")]
        [Authorize]
        public async Task<IActionResult> UpdateDataInput([FromBody] InsertModel model)
        {
            var transferObject = new TransferObject();

            await _service.UpdateDataInput(model);
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
                transferObject.GetMessage("0101", _service);
            }
            return Ok(transferObject);
        }

        //[HttpPost("ExportExcel")]
        ////[Authorize]
        //public async Task<IActionResult> ExportExcel([FromBody] CalculateResultModel data, [FromQuery] string headerId)
        //{
        //    var transferObject = new TransferObject();
        //    MemoryStream outFileStream = new MemoryStream();
        //    var path = Directory.GetCurrentDirectory() + "/Template/CoSoTinhMucGiamGia.xlsx";
        //    var k = await _service.ExportExcel(outFileStream, path, headerId, data);
        //    if (_service.Status)
        //    {
        //        var result = await _service.SaveFileHistory(k, headerId);
        //        //return File(outFileStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", DateTime.Now.ToString() + "_CoSoTinhMucGiamGia" + ".xlsx");
        //        transferObject.Data = result;
        //        return Ok(transferObject);
        //    }
        //    else
        //    {
        //        transferObject.Status = false;
        //        transferObject.MessageObject.MessageType = MessageType.Error;
        //        transferObject.GetMessage("2000", _service);
        //        return Ok(transferObject);
        //    }
        //}

        [HttpPost("ExportExcel")]
        public async Task<IActionResult> ExportExcel([FromQuery] string headerId, [FromBody] CalculateResultModel data)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile(new List<string>(), "EXCEL", headerId, data);
            if (_service.Status)
            {
                transferObject.Data = result;
                return Ok(transferObject);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }

        [HttpPost("ExportWord")]
        [Authorize]
        public async Task<IActionResult> ExportWord([FromBody] List<string> lstCustomerChecked, [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile(lstCustomerChecked, "WORD", headerId, new CalculateResultModel());
            if (_service.Status)
            {
                transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }

        [HttpPost("ExportWordTrinhKy")]
        [Authorize]
        public async Task<IActionResult> ExportWordTrinhky([FromBody] List<string> lstTrinhKyChecked, [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile(lstTrinhKyChecked, "WORDTRINHKY", headerId, new CalculateResultModel());
            if (_service.Status)
            {
                transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }

        [HttpPost("ExportPDF")]
        [Authorize]
        public async Task<IActionResult> ExportPDF([FromBody] List<string> lstCustomerChecked, [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile(lstCustomerChecked, "PDF", headerId, new CalculateResultModel());
            if (_service.Status)
            {
                transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }
        [HttpGet("SendMail")]
        [Authorize]
        public async Task<IActionResult> SendMail( [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
             await _service.SendEmail(headerId);
            if (_service.Status)
            {
                //transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }
        [HttpGet("SendSMS")]
        [Authorize]
        public async Task<IActionResult> SendSMS([FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            await _service.SendSMS(headerId);
            if (_service.Status)
            {
                //transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }
        [HttpGet("Getmail")]
        [Authorize]
        public async Task<IActionResult> Getmail([FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result= await _service.GetMail(headerId);
            if (_service.Status)
            {
                transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }
        [HttpGet("GetSms")]
        [Authorize]
        public async Task<IActionResult> GetSms([FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetSms(headerId);
            if (_service.Status)
            {
                transferObject.Data = result;
                return Ok(transferObject);

            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }
    }
}
