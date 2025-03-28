using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.BU;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    [Authorize]
    public class ListTablesController : ControllerBase
    {
        public readonly IListTablesService _service;
        public ListTablesController(IListTablesService service)
        {
            _service = service;
        }
        [HttpGet("GetListTablesTree")]
        public async Task<IActionResult> GetListTablesTree()
        {
            var transferObject = new TransferObject();
            var result = await _service.BuildDataForTree();
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
        [HttpGet("GetListTablesTreeWithMgCode/{mgCode}")]
        public async Task<IActionResult> GetListTablesTreeWithMgCode([FromRoute] string mgCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.BuildDataForTreeWithMgCode(mgCode);
            if (_service.Status && result != null)
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

        [HttpGet("GetAll")]
        public async Task<IActionResult> GetAll([FromQuery] BaseFilter filter)
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


        [HttpPut("Update-Order")]
        public async Task<IActionResult> UpdateOrganize([FromBody] ListTablesDto moduleDto)
        {
            var transferObject = new TransferObject();
            await _service.UpdateOrderTree(moduleDto);
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

        [HttpPost("Insert")]
        public async Task<IActionResult> Insert([FromBody] ListTablesDto tables)
        {
            var transferObject = new TransferObject();
            var result = await _service.Add(tables);
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
        [HttpPut("UploadExcel")]
        public async Task<IActionResult> UploadExcel(IFormFile file, [FromQuery] string mgCode)
        {
            var transferObject = new TransferObject();
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var result = await _service.ImportExcel(file, mgCode);
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
        [HttpPut("Update")]
        public async Task<IActionResult> Update([FromBody] ListTablesDto moduleDto)
        {
            var transferObject = new TransferObject();
            await _service.Update(moduleDto);
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
        public async Task<IActionResult> Delete([FromRoute] Guid code)
        {
            var transferObject = new TransferObject();

            var result = await _service.Delete(code);
            if (_service.Status && result != null)
            {
                transferObject.Status = true;
                transferObject.Data = result;
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


    }
}
