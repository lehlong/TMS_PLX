using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Dtos.AD;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.BU;
using DMS.CORE.Entities.MD;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    public class OpinionListController : ControllerBase
    {
        public readonly IOpinionListService _service;
        public OpinionListController(IOpinionListService service)
        {
            _service = service;
        }
        [HttpGet("GetOpinionListTree")]
        public async Task<IActionResult> GetOpinionListTree()
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
                transferObject.GetMessage("2000", _service);
            }
            return Ok(transferObject);
        }
        [HttpGet("GetOpinionListTreeWithMgCode/{mgCode}")]
        public async Task<IActionResult> GetOpinionListTreeWithMgCode([FromRoute] string mgCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.BuildDataForTreeWithMgCode(mgCode);
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

        [HttpGet("GetOplTreeWithMgCodeAndOrg/{orgCode}/{mgCode}")]
        public async Task<IActionResult> GetOplTreeWithMgCodeAndOrg([FromRoute] string mgCode, [FromRoute] string orgCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.BuildDataForTreeWithMgCodeAndOrg(mgCode, orgCode);
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
        [HttpGet("GetOrgInOpinion/{Code}")]
        public async Task<IActionResult> GetTemplateWithTreeOrganize(Guid Code)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetOpinionListWithTreeOrganize(Code);
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

        [HttpGet("GetOpinionDetail/{MgCode}/{OrgCode}/{OpinionCode}")]
        public async Task<IActionResult> GetOpinionDetail([FromRoute] string MgCode, [FromRoute] string OrgCode, [FromRoute] string OpinionCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetOpinionDetail(MgCode, OrgCode, OpinionCode);
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

        [HttpPut("UpdateOpinionDetail")]
        public async Task<IActionResult> UpdateOpinionDetail([FromBody] OpinionDetailDto opd)
        {
            var transferObject = new TransferObject();
            var data = await _service.UpdateOpinionDetail(opd);
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


        [HttpGet("GetOpinionListTreeWithTimeYearAndAuditPeriod")]
        public async Task<IActionResult> GetOpinionListTreeWithTimeYearAndAuditPeriod([FromQuery] string timeYear, [FromQuery] string auditPeriod)
        {
            var transferObject = new TransferObject();
            var result = await _service.BuildDataForTreeWithTimeYearAndAuditPeriod(timeYear, auditPeriod);
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
        public async Task<IActionResult> UpdateOrganize([FromBody] OpinionListDto moduleDto)
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
        public async Task<IActionResult> Insert([FromBody] OpinionListDto opinion)
        {
            var transferObject = new TransferObject();
            var result = await _service.Add(opinion);
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
        public async Task<IActionResult> Update([FromBody] OpinionListUpdateDto moduleDto)
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
    }
}
