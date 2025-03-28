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
    public class AuditPeriodListTablesController(IAuditPeriodListTablesService service) : ControllerBase
    {
        public readonly IAuditPeriodListTablesService _service = service;
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
        public async Task<IActionResult> Insert([FromBody] AuditPeriodListTablesCreateDto dto)
        {
            var transferObject = new TransferObject();
            var result = await _service.Add(dto);
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
        [HttpGet("GetHistoryByListAuditCode/{listAuditCode}")]
        public async Task<IActionResult> GetHistoryByListAuditCode([FromRoute] string listAuditCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetHistoryByListAuditCode(listAuditCode);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2001", _service);
            }
            return Ok(transferObject);
        }
        [HttpPut("Update")]
        public async Task<IActionResult> Update([FromBody] AuditPeriodListTablesDto time)
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
        [HttpPut("UpdateMultipleAuditTemplateListTablesData")]
        public async Task<IActionResult> UpdateMultipleAuditTemplateListTablesData([FromBody] List<AuditTemplateListTablesDataDto> dtos)
        {
            var transferObject = new TransferObject();
            foreach (var dto in dtos)
            {
                await _service.UpdateAuditTemplateListTablesData(dto);
                if (!_service.Status)
                {
                    transferObject.Status = false;
                    transferObject.MessageObject.MessageType = MessageType.Error;
                    transferObject.GetMessage("0104", _service);
                    return Ok(transferObject); // Trả về ngay nếu có lỗi
                }
            }
            transferObject.Status = true;
            transferObject.MessageObject.MessageType = MessageType.Success;
            transferObject.GetMessage("0103", _service);
            return Ok(transferObject);
        }
        [HttpDelete("Delete/{auditPeriodCode}")]
        public async Task<IActionResult> Delete([FromRoute] int auditPeriodCode)
        {
            var transferObject = new TransferObject();
            await _service.Delete(auditPeriodCode);
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
        [HttpGet("GetMgListTablesByAuditPeriodCode/{auditPeriodCode}")]
        public async Task<IActionResult> GetMgListTablesByAuditPeriodCode([FromRoute] string auditPeriodCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetMgListTablesByAuditPeriodCode(auditPeriodCode);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2001", _service);
            }
            return Ok(transferObject);
        }
        [HttpGet("GetTemplateListTablesByAuditPeriodCode")]
        public async Task<IActionResult> GetMgListTablesByAuditPeriodCode([FromQuery] string auditPeriodCode, [FromQuery] Guid templateGroupsCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetTemplateListTablesByAuditPeriodCode(auditPeriodCode, templateGroupsCode);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2001", _service);
            }
            return Ok(transferObject);
        }
        [HttpGet("GetTemDataWithMgListTablesAndOrgCode")]
        public async Task<IActionResult> GetTemDataWithMgListTablesAndOrgCode([FromQuery] int auditListTablesCode, [FromQuery] string? orgCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.GetTemDataWithMgListTablesAndOrgCode(auditListTablesCode, orgCode);
            if (_service.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2001", _service);
            }
            return Ok(transferObject);
        }
        [HttpPut("ChangeStatusReview/{code}")]
        public async Task<IActionResult> ChangeStatusReview([FromRoute] int code, [FromQuery] string? textContent)
        {
            var transferObject = new TransferObject();
            var result = await _service.ChangeStatusReview(code, textContent);
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
        [HttpPut("ChangeStatusApproval/{code}")]
        public async Task<IActionResult> ChangeStatusApproval([FromRoute] int code)
        {
            var transferObject = new TransferObject();
            var result = await _service.ChangeStatusApproval(code);
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
        [HttpPut("ChangeStatusCancel/{code}")]
        public async Task<IActionResult> ChangeStatusCancel([FromRoute] int code, [FromQuery] string? action,
            [FromQuery] string? textContent)
        {
            var transferObject = new TransferObject();
            var result = await _service.ChangeStatusCancel(code, action, textContent);
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
        [HttpPut("ChangeStatusConfirm/{code}")]
        public async Task<IActionResult> ChangeStatusConfirm([FromRoute] int code, [FromQuery] string? action,
            [FromQuery] string? textContent)
        {
            var transferObject = new TransferObject();
            var result = await _service.ChangeStatusconfirm(code, action, textContent);
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

        [HttpPut("UploadExcelSTC")]
        public async Task<IActionResult> UploadExcelSTC(IFormFile file, [FromQuery] int auditListTablesCode)
        {
            var transferObject = new TransferObject();
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var result = await _service.ImportExcelAndUpdateDataSTC(file, auditListTablesCode);
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
        [HttpPut("UploadExcelDV")]
        public async Task<IActionResult> UploadExcelDV(IFormFile file, [FromQuery] int auditListTablesCode)
        {
            var transferObject = new TransferObject();
            if (file == null || file.Length == 0)
            {
                return BadRequest("No file uploaded.");
            }

            var result = await _service.ImportExcelAndUpdateDataDV(file, auditListTablesCode);
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
        public async Task<IActionResult> Export([FromQuery] int auditListTablesCode, [FromQuery] string? orgCode)
        {
            var transferObject = new TransferObject();
            var result = await _service.ExportAuditTemplateListTablesData(auditListTablesCode, orgCode);
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
    }
}
