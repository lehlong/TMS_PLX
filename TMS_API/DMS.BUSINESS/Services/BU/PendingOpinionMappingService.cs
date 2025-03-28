using AutoMapper;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.BU;
using DMS.CORE;
using DMS.CORE.Entities.BU;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.BU
{
    public interface IPendingOpinionMappingService : IGenericService<tblBuPendingOpinionMapping, PendingOpinionMappingDto>
    {
    }
    public class PendingOpinionMappingService(AppDbContext dbContext, IMapper mapper) : GenericService<tblBuPendingOpinionMapping, PendingOpinionMappingDto>(dbContext, mapper), IPendingOpinionMappingService
    {

    }
}
