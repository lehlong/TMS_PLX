using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.CORE;
using DMS.CORE.Entities.BU;

namespace DMS.BUSINESS.Services.BU
{
    public interface IExportDataService : IGenericService<TblBuCalculateDiscount, CalculateDiscountDto>
    {
      
    }
    public class ExportDataService(AppDbContext dbContext, IMapper mapper) : GenericService<TblBuCalculateDiscount, CalculateDiscountDto>(dbContext, mapper), IExportDataService
    {
    }
}
