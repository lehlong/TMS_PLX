using Microsoft.AspNetCore.Mvc;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using Microsoft.AspNetCore.StaticFiles;
using DMS.BUSINESS.Dtos.BU;
using DMS.BUSINESS.Services.BU.Attachment;
using Common;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    public class AttachmentController(IAttachmentService service) : ControllerBase
    {
        public readonly IAttachmentService _service = service;

        [HttpDelete("Delete/{Id}")]
        public async Task<IActionResult> Delete([FromRoute] int Id)
        {
            var transferObject = new TransferObject();
            await _service.Delete(Id);
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

        [HttpGet("Download")]
        public async Task<IActionResult> Download([FromQuery] Guid attachmentId)
        {
            var transferObject = new TransferObject();
            var result = await _service.Download(attachmentId);
            var contentTypeProvider = new FileExtensionContentTypeProvider();
            contentTypeProvider.TryGetContentType(result.Item2, out string contentType);
            if (_service.Status)
            {
                return File(result.Item1, contentType, result.Item2);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("2000", _service);
                return Ok(transferObject);
            }

        }

        [HttpPut("Update")]
        public async Task<IActionResult> Update([FromBody] AttachmentRenameDto dto)
        {
            var transferObject = new TransferObject();
            await _service.Update(dto);
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
    }
}
