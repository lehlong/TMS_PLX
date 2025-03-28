using Common;
using DMS.API.AppCode.Enum;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.MD;
using Microsoft.AspNetCore.Mvc;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Models;

namespace DMS.API.Controllers.MD
{
    [Route("api/[controller]")]
    [ApiController]
    public class CustomerContactController(ICustomerContactService service) : ControllerBase
    {
        public readonly ICustomerContactService _service = service;
        [HttpGet("GetAll")]
    public async Task<IActionResult> GetAll([FromQuery] BaseMdFilter filter)
    {
        var transferObject = new TransferObject();
        var result = await _service.GetAll(filter);
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
        [HttpPost("Insert")]
        public async Task<IActionResult> Insert([FromBody] CustomerContactModel contact)
        {
            var transferObject = new TransferObject();

            var result = await _service.Insert(contact); // Thêm nhiều bản ghi cùng lúc

            if (_service.Status)
            {
                transferObject.Data = result;
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
        [HttpPut("Update")]
        public async Task<IActionResult> Update([FromBody] CustomerContactModel contact)
        {
            var transferObject = new TransferObject();

            var result = await _service.UpdateCustomerContact(contact);

            if (_service.Status)
            {
                transferObject.Data = result;
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
    }
}
