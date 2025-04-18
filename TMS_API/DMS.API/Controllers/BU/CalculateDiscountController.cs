using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Models;
using DMS.BUSINESS.Services.BU;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    public class CalculateDiscountController(ICalculateDiscountService service) : ControllerBase
    {
        public readonly ICalculateDiscountService _service = service;

        [HttpGet("Search")]
        public async Task<IActionResult> Search([FromQuery] BaseFilter filter)
        {
            var transferObject = new TransferObject();
            var result = await _service.Search(filter);
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

        [HttpGet("GenarateCreate")]
        public async Task<IActionResult> GenarateCreate()
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateCreate();
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

        [HttpPost("Create")]
        public async Task<IActionResult> Create([FromBody] CalculateDiscountInputModel input)
        {
            var transferObject = new TransferObject();
            await _service.Create(input);
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

        [HttpGet("GetInput")]
        public async Task<IActionResult> GetInput([FromQuery] string id)
        {
            var transferObject = new TransferObject();
            var data = await _service.GetInput(id);
            if (_service.Status)
            {
                transferObject.Data = data;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
            }
            return Ok(transferObject);
        }

        [HttpGet("CopyInput")]
        public async Task<IActionResult> CopyInput([FromQuery] string headerId, string id)
        {
            var transferObject = new TransferObject();
            var data = await _service.CopyInput(headerId, id);
            if (_service.Status)
            {
                transferObject.Data = data;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
            }
            return Ok(transferObject);
        }

        [HttpPut("UpdateInput")]
        public async Task<IActionResult> UpdateInput([FromBody] CalculateDiscountInputModel input)
        {
            var transferObject = new TransferObject();
            await _service.UpdateInput(input);
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

        [HttpPut("HandleQuyTrinh")]
        public async Task<IActionResult> HandleQuyTrinh([FromBody] QuyTrinhModel data)
        {
            var transferObject = new TransferObject();
            await _service.HandleQuyTrinh(data);
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

        [HttpGet("GetOutput")]
        public async Task<IActionResult> GetOutput([FromQuery] string id)
        {
            var transferObject = new TransferObject();
            var data = await _service.CalculateDiscountOutput(id);
            if (_service.Status)
            {
                transferObject.Data = data;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
            }
            return Ok(transferObject);
        }

        [HttpGet("ExportExcel")]
        public async Task<IActionResult> ExportExcel([FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.ExportExcel(headerId);
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
        public async Task<IActionResult> ExportWordTrinhky([FromBody] List<string> lstTrinhKyChecked, [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile(lstTrinhKyChecked, "WORDTRINHKY", headerId, new CalculateDiscountInputModel(), []);
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

        [HttpGet("SendMail")]
        [Authorize]
        public async Task<IActionResult> SendMail([FromQuery] string headerId)
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
        public async Task<IActionResult> SendSMS([FromQuery] string headerId, string smsName)
        {
            var transferObject = new TransferObject();
            await _service.SendSMS(headerId, smsName);
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
            var result = await _service.GetHistoryMail(headerId);
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
            var result = await _service.GetHistorySms(headerId);
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
        [HttpPost("ExportWord")]
        [Authorize]
        public async Task<IActionResult> ExportWord([FromBody] List<CustomBBDOExportWord> lstCustomerChecked, [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile([], "WORD", headerId, new CalculateDiscountInputModel(), lstCustomerChecked);
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
        public async Task<IActionResult> ExportPDF([FromBody] List<CustomBBDOExportWord> lstCustomerChecked, [FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            var result = await _service.GenarateFile([], "PDF", headerId, new CalculateDiscountInputModel(), lstCustomerChecked);
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

        [HttpGet("GetCustomerBbdo")]
        [Authorize]
        public async Task<IActionResult> GetCustomerBBDO([FromQuery] string id)
        {
            var transferObject = new TransferObject();
            var data = await _service.GetCustomerBbdo(id);

            if (_service.Status)
            {
                transferObject.Data = data;
                return Ok(transferObject);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                return BadRequest(transferObject);
            }
        }
        [HttpGet("GetAllInputCustomer")]
        [Authorize]
        public async Task<IActionResult> GetAllInputCustomer()
        {
            var transferObject = new TransferObject();
            var data = await _service.GetAllInputCustomer();

            if (_service.Status)
            {
                transferObject.Data = data;
                return Ok(transferObject);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                return BadRequest(transferObject);
            }
        }
    }
}