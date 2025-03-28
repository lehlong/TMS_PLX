using Microsoft.AspNetCore.Http;
using System.ComponentModel.DataAnnotations;

namespace DMS.BUSINESS.Dtos.MD
{
    public class CuocVanChuyenListImportDto
    {
        [Required]
        public IFormFile File { get; set; }

        [Required]
        public string HeaderCode { get; set; }
    }
}
