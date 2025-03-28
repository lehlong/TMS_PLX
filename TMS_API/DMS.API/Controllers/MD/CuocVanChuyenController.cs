using Microsoft.AspNetCore.Mvc;
using DMS.BUSINESS.Services.MD;
using DMS.BUSINESS.Dtos.MD;
using Common;
using DMS.CORE;
using AutoMapper;
using DMS.API.AppCode.Enum;
using DMS.API.AppCode.Extensions;

namespace DMS.API.Controllers.MD
{
    [ApiController]
    [Route("api/[controller]")]
    public class CuocVanChuyenController : ControllerBase
    {
        private readonly ICuocVanChuyenService _cuocVanChuyenService;
        private readonly IMapper _mapper;
        private readonly AppDbContext _dbContext;
        public CuocVanChuyenController(ICuocVanChuyenService cuocVanChuyenService, IMapper mapper, AppDbContext dbContext)
        {
            _cuocVanChuyenService = cuocVanChuyenService;
            _mapper = mapper;
            _dbContext = dbContext;
        }

        [HttpGet("SearchById/{id}")]
        public async Task<IActionResult> SearchById([FromQuery] BaseFilter filter, string id)
        {
            var transferObject = new TransferObject();
            var result = await _cuocVanChuyenService.SearchById(filter, id);
            if (_cuocVanChuyenService.Status)
            {
                transferObject.Data = result;
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0001", _cuocVanChuyenService);
            }
            return Ok(transferObject);
        }

        [HttpGet("GetAll")]
        public async Task<IActionResult> GetAll([FromQuery] BaseMdFilter filter)
        {
            var result = await _cuocVanChuyenService.GetAll(filter);
            if (result == null)
            {
                return BadRequest("Không tìm thấy dữ liệu");
            }
            return Ok(result);
        }

        [HttpPost("importExcel")]
        public async Task<IActionResult> ImportExcelData([FromForm] CuocVanChuyenListImportDto data)
        {
            var transferObject = new TransferObject();
            if (data.File == null || data.File.Length == 0)
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0101", _cuocVanChuyenService);
                return Ok(transferObject);
            }

            var fileExtension = Path.GetExtension(data.File.FileName).ToLower();
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0102", _cuocVanChuyenService);
                return Ok(transferObject);
            }

            string filePath = Path.GetTempFileName();

            using (var memoryStream = new MemoryStream())
            {
                await data.File.CopyToAsync(memoryStream);
                await System.IO.File.WriteAllBytesAsync(filePath, memoryStream.ToArray());
            }

            var result = await _cuocVanChuyenService.ImportExcelData(filePath, data.HeaderCode);

            System.IO.File.Delete(filePath);

            if (result != null)
            {
                transferObject.Data = result;
                transferObject.Status = true;
                transferObject.MessageObject.MessageType = MessageType.Success;
                transferObject.GetMessage("0100", _cuocVanChuyenService);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0101", _cuocVanChuyenService);
            }
            return Ok(transferObject);
        }


        [HttpPost("Insert")]
        public async Task<IActionResult> Insert([FromBody] CuocVanChuyenDto CuocVanChuyen)
        {
            var transferObject = new TransferObject();
            CuocVanChuyen.Code = Guid.NewGuid().ToString();
            var result = await _cuocVanChuyenService.Add(CuocVanChuyen);
            if (_cuocVanChuyenService.Status)
            {
                transferObject.Data = result;
                transferObject.Status = true;
                transferObject.MessageObject.MessageType = MessageType.Success;
                transferObject.GetMessage("0100", _cuocVanChuyenService);
            }
            else
            {
                transferObject.Status = false;
                transferObject.MessageObject.MessageType = MessageType.Error;
                transferObject.GetMessage("0101", _cuocVanChuyenService);
            }
            return Ok(transferObject);
        }
    }
}
