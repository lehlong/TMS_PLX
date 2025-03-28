using AutoMapper;
using Common;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.MD;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Dtos.BU
{
    public class ListTablesDto : IMapFrom, IDto
    {

        [Key]
        public Guid Code { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string PId { get; set; }
        public int? OrderNumber { get; set; }

        public bool? IsChecked { get; set; }
        public bool? IsLeaf { get; set; }

        public string? Title { get; set; }

        public string? Key { get; set; }
        public string? MgCode { get; set; }
        public string? CurrencyCode { get; set; }
        public TblMdCurrency? Currency { get; set; }
        public TblMdMgListTables? ListTables { get; set; }
        public List<ListTablesDto>? Children { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblBuListTables, ListTablesDto>().ReverseMap();
        }
    }
    public class ListTablesViewModel
    {
        public Guid GroupId { get; set; }

        public List<ListTablesDto> ListTables { get; set; }

        public ListTablesViewModel()
        {
            ListTables = new List<ListTablesDto>();
        }
    }
    public class ListTablesUpdateDto : IMapFrom, IDto
    {

        [Key]
        public Guid Code { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string PId { get; set; }
        public int? OrderNumber { get; set; }
        public string? MgCode { get; set; }
        public string? CurrencyCode { get; set; }
        public List<ListTablesDto>? Children { get; set; }
        public void Mapping(Profile profile)
        {
            profile.CreateMap<TblBuListTables, ListTablesUpdateDto>().ReverseMap();
        }
    }
}
