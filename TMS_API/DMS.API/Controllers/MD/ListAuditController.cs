using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Services.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Http.HttpResults;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.MD
{
    [Route("api/[controller]")]
    [ApiController]
    public class ListAuditController(IListAuditService service) : ControllerBase
    {

        public readonly IListAuditService _service = service;
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
        [HttpGet("GetByCode/{code}")]
        public async Task<IActionResult> GetByCode([FromRoute] string code)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetByCodeAsync(code);
            if (_service.Status && result != null)
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
        [HttpGet("GetListAuditHistory/{auditCode}")]
        public async Task<IActionResult> GetListAuditHistory([FromRoute] string auditCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetListAuditHistory(auditCode);
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
        [HttpGet("GetOpinionListUnfinished")]
        public async Task<IActionResult> GetOpinionListUnfinished([FromQuery] string timeYear, [FromQuery] string auditPeriod)
        {
            var transferObject = new TransferObject();
            var result = await _service.BuildDataForTreeUnfinished(timeYear, auditPeriod);
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
        [HttpPut("UpdateListAudit")]
        public async Task<IActionResult> UpdateListAudit([FromBody] ListAuditDto audit)
        {
            var transferObject = new TransferObject();
            var data = await _service.UpdateListAudit(audit);
            if (_service.Status)
            {
                transferObject.Data = data;
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
        public async Task<IActionResult> Insert([FromBody] ListAuditDto audit)
        {
            var transferObject = new TransferObject();
            var result = await _service.Add(audit);
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
        public async Task<IActionResult> Update([FromBody] ListAuditDto audit)
        {
            var transferObject = new TransferObject();
            await _service.Update(audit);
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
        [HttpGet("getOpinionStatis/{auditCode}")]
        public async Task<IActionResult> getOpinionStatis([FromRoute] string auditCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.getOpinionStatis(auditCode);
            if (_service.Status)
            {
                transferObject.Data = result;
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
        [HttpGet("GetListOrg/{auditCode}")]
        public async Task<IActionResult> GetListOrg([FromRoute] string auditCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetOrganizeListAsync(auditCode);
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
