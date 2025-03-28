using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.BU;
using DocumentFormat.OpenXml.Math;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class DiscountInformationListController(IDiscountInformationListService service) : ControllerBase
    {
        public readonly IDiscountInformationListService _service = service;
        
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
        public async Task<IActionResult> Insert([FromBody] CompetitorModel model)
        {
            var transferObject = new TransferObject();

            await _service.InsertData(model);
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

        [HttpGet("GetObjectCreate")]
        public async Task<IActionResult> GetObjectCreate(string code)
        {
            var i = code;
            var transferObject = new TransferObject();
            var result = await _service.BuildObjectCreate(code);
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

        //[HttpGet("GetDataInput")]
        //public async Task<IActionResult> GetDataInput(string code)
        //{
        //    var transferObject = new TransferObject();
        //    var result = await _service.BuidDataInput(code);

        //    return null;
        //}
    }
}
