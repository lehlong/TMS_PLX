using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Models;
using DMS.BUSINESS.Services.BU;
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
    }
}
