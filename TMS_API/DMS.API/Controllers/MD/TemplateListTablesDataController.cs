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
    public class TemplateListTablesDataController(ITemplateListTablesDataService service) : ControllerBase
    {
        public readonly ITemplateListTablesDataService _service = service;
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
        public async Task<IActionResult> Insert([FromBody] List<TemplateListTablesDataDto> opinion)
        {
            var transferObject = new TransferObject();
            var results = new List<TemplateListTablesDataDto>(); // Danh sách lưu kết quả

            foreach (var item in opinion)
            {
                var result = await _service.Add(item);
                if (result != null)
                {
                    results.Add(result);
                }
            }

            if (_service.Status)
            {
                transferObject.Data = results;
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
        public async Task<IActionResult> Update([FromBody] List<TemplateListTablesDataDto> dtos)
        {
            var transferObject = new TransferObject();
            var result = await _service.Update(dtos);
            if (result)
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
        [HttpDelete("Delete")]
        public async Task<IActionResult> Delete([FromQuery] TemplateListTablesDataGenCodeDto dto)
        {
            var transferObject = new TransferObject();
            await _service.Delete(dto);
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
        [HttpGet("export")]
        public async Task<IActionResult> Export([FromQuery] string templateCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.Export(templateCode);
            if (_service.Status)
            {
                return File(result, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "DSTLTB" + DateTime.Now.ToString() + ".xlsx");
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }
        }
        //[HttpGet("GetTemplateWithTree/{id}")]
        //public async Task<IActionResult> GetTemplateWithTreeOrganize(string id)
        //{
        //    var transferObject = new TransferObject();
        //    var result = await _service.GetTemplateWithTree(id);
        //    if (_service.Status)
        //    {
        //        transferObject.Data = result;
        //    }
        //    else
        //    {
        //        transferObject.Status = false;
        //        transferObject.MessageObject.MessageType = MessageType.Error;
        //        transferObject.GetMessage("0001", _service);
        //    }
        //    return Ok(transferObject);
        //}
    }
}
