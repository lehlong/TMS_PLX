using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.MD
{
    [Route("api/[controller]")]
    [ApiController]
    public class PeriodTimeController(IPeriodTimeService service) : ControllerBase
    {
        public readonly IPeriodTimeService _service = service;
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
        public async Task<IActionResult> Insert([FromBody] PeriodTimeDto time)
        {
            var transferObject = new TransferObject();
            var result = await _service.Add(time);
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
        public async Task<IActionResult> Update([FromBody] PeriodTimeDto time)
        {
            var transferObject = new TransferObject();
            await _service.Update(time);
            if (_service.Status)
            {
                transferObject.Status = true;
                transferObject.MessageObject.MessageType = MessageType.Success;
                transferObject.GetMessage("0103", _service);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0104", _service);
            }
            return Ok(transferObject);
        }
        [HttpDelete("Delete/{code}")]
        public async Task<IActionResult> Delete([FromRoute] string code)
        {
            var transferObject = new TransferObject();
            await _service.Delete(code);
            if (_service.Status)
            {
                transferObject.Status = true;
                transferObject.MessageObject.MessageType = MessageType.Success;
                transferObject.GetMessage("0105", _service);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0106", _service);
            }
            return Ok(transferObject);
        }
        [HttpPut("ChangeDefault/{timeYear}")]
        public async Task<IActionResult> ChangeDefault([FromRoute] int timeYear)
        {
            var transferObject = new TransferObject();
            var result = await _service.ChangeDefaultStatus(timeYear);
            if (result)
            {
                transferObject.Status = true;
                transferObject.MessageObject.MessageType = MessageType.Success;
                transferObject.GetMessage("0105", _service); //"0107" là mã thông báo cho thay đổi trạng thái mặc định thành công
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0106", _service); // "0108" là mã thông báo cho thay đổi trạng thái mặc định thất bại
            }
            return Ok(transferObject);
        }
        [HttpPut("ChangeIsClosed/{timeYear}")]
        public async Task<IActionResult> ChangeIsClosed([FromRoute] int timeYear)
        {
            var transferObject = new TransferObject();
            try
            {
                var result = await _service.ChangeIsClosedStatus(timeYear);
                if (result)
                {
                    transferObject.Status = true;
                    transferObject.MessageObject.MessageType = MessageType.Success;
                    transferObject.GetMessage("0105", _service); // Giả sử "0109" là mã thông báo cho thay đổi trạng thái IsClosed thành công
                }
                else
                {
                    transferObject.Status = false;
                    transferObject.MessageObject.MessageType = MessageType.Error;
                    transferObject.GetMessage("0106", _service); // Giả sử "0110" là mã thông báo cho thay đổi trạng thái IsClosed thất bại
                }
            }
            catch (Exception ex)
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.MessageObject.Message = "An error occurred while changing IsClosed status.";
                transferObject.MessageObject.MessageDetail = ex.Message;
            }
            return Ok(transferObject);
        }

    }
}
