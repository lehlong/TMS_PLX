using Common;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;
using DMS.BUSINESS.Services.BU;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;

namespace DMS.API.Controllers.BU
{
    [Route("api/[controller]")]
    [ApiController]
    public class DiscountInformationController : ControllerBase
    {
        public readonly IDiscountInformationService _service;

        public DiscountInformationController(IDiscountInformationService service)
        {
            _service = service;
        }   

        [HttpGet("GetAll")]
        [Authorize]
        public async Task<IActionResult> GetAll([FromQuery] string Code)
        {
            var transferObject = new TransferObject();
            var result = await _service.getAll(Code);
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

        //[HttpGet("GetDataInput")]
        //public async Task<IActionResult> GetDataInput([FromQuery] string code)
        //{
        //    var transferObject = new TransferObject();
        //    var result = await _service.GetDataInput(code);
        //    if (_service.Status)
        //    {
        //        transferObject.Data = result;
        //    }
        //    else
        //    {
        //        transferObject.Status = false;
        //        transferObject.MessageObject.MessageType = MessageType.Error;
        //        //transferObject.GetMessage("2000", _service);
        //    }
        //    return Ok(transferObject);
        //}

        [HttpPost("UpdateDataInput")]
        [Authorize]
        public async Task<IActionResult> UpdateDataInput([FromBody] CompetitorModel model)
        {
            var transferObject = new TransferObject();

            await _service.UpdateDataInput(model);
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

        [HttpGet("ExportExcel")]
        [Authorize]
        public async Task<IActionResult> ExportExcel([FromQuery] string headerId)
        {
            var transferObject = new TransferObject();
            MemoryStream outFileStream = new MemoryStream();
            var path = Directory.GetCurrentDirectory() + "/Template/PhanTichChietKhau.xlsx";
            _service.ExportExcel(ref outFileStream, path, headerId);
            if (_service.Status)
            {
                var result = await _service.SaveFileHistory(outFileStream, headerId);
                //return File(outFileStream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", DateTime.Now.ToString() + "_CoSoTinhMucGiamGia" + ".xlsx");
                transferObject.Data = result;
                return Ok(transferObject);
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
