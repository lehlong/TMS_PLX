using Microsoft.AspNetCore.Mvc;
using DMS.BUSINESS.Services.MD;
using DMS.BUSINESS.Dtos.MD;
using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;

namespace DMS.API.Controllers.MD
{
    [ApiController]
    [Route("api/[controller]")]
    public class CuocVanChuyenListController(ICuocVanChuyenListService service) : ControllerBase
    {
        public readonly ICuocVanChuyenListService _service = service;
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
        public async Task<IActionResult> Insert([FromBody] CuocVanChuyenListDto CuocVanChuyen)
        {
            var transferObject = new TransferObject();
            CuocVanChuyen.Code = Guid.NewGuid().ToString();
            var result = await _service.Add(CuocVanChuyen);
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
