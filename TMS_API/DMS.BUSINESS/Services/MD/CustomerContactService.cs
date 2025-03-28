using AutoMapper;
using AutoMapper.Internal;
using Common;
using DMS.BUSINESS.Common;
using DMS.BUSINESS.Dtos.MD;
using DMS.BUSINESS.Models;
using DMS.CORE;
using DMS.CORE.Entities.MD;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.BUSINESS.Services.MD
{
    public interface ICustomerContactService : IGenericService<TblMdCustomerContact, CustomerContactDto>
    {
        Task<IList<CustomerContactDto>> GetAll(BaseMdFilter filter);
        Task<IList<CustomerContactDto>> Insert(CustomerContactModel contact);
        Task<IList<CustomerContactDto>> UpdateCustomerContact(CustomerContactModel contact);
    }
   public class CustomerContactService(AppDbContext dbContext, IMapper mapper) : GenericService<TblMdCustomerContact, CustomerContactDto>(dbContext, mapper), ICustomerContactService
    {
        public async Task<IList<CustomerContactDto>> GetAll(BaseMdFilter filter)
        {
            try
            {
                var query = _dbContext.TblMdCustomerContact.AsQueryable();
                if (filter.IsActive.HasValue)
                {
                    query = query.Where(x => x.IsActive == filter.IsActive);
                }
                return await base.GetAllMd(query, filter);
            }
            catch (Exception ex)
            {
                Status = false;
                Exception = ex;
                return null;
            }
        }
        public async Task<IList<CustomerContactDto>> Insert(CustomerContactModel contact)
        {
            using var transaction = await _dbContext.Database.BeginTransactionAsync();
            try
            {
                var entities = contact.Contact_List.Select(c => new TblMdCustomerContact
                {
                    Code = Guid.NewGuid().ToString(),
                    Customer_Code = contact.Customer_Code, // Gán đúng mã khách hàng
                    Type = c.Type,
                    Value = c.Value,
                    IsActive = true
                }).ToList();

                await _dbContext.TblMdCustomerContact.AddRangeAsync(entities); // Thêm nhiều bản ghi
                await _dbContext.SaveChangesAsync(); // Lưu vào database

                await transaction.CommitAsync(); // Commit transaction

                return _mapper.Map<List<CustomerContactDto>>(entities); // Trả về danh sách đã thêm
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync(); // Rollback nếu có lỗi
                Status = false;
                Exception = ex;
                return new List<CustomerContactDto>();
            }
        }

        public async Task<IList<CustomerContactDto>> UpdateCustomerContact(CustomerContactModel contactList)
        {
            using var transaction = await _dbContext.Database.BeginTransactionAsync();
            try
            {
                // Chia dữ liệu thành 2 nhóm: cần Insert và cần Update
                var newContacts = contactList.Contact_List
                    .Where(c => string.IsNullOrEmpty(c.Code)) // Những bản ghi mới cần insert
                    .Select(c => new TblMdCustomerContact
                    {
                        Code = Guid.NewGuid().ToString(),
                        Customer_Code = contactList.Customer_Code,
                        Type = c.Type,
                        Value = c.Value,
                        IsActive = true
                    })
                    .ToList();

                var existingContacts = contactList.Contact_List
                    .Where(c => !string.IsNullOrEmpty(c.Code)) // Những bản ghi cần update
                    .ToList();

                // Thêm mới
                if (newContacts.Any())
                {
                    await _dbContext.TblMdCustomerContact.AddRangeAsync(newContacts);
                }

                // Cập nhật danh sách cũ
                if (existingContacts.Any())
                {
                    var contactCodes = existingContacts.Select(c => c.Code).ToList();
                    var contactsToUpdate = await _dbContext.TblMdCustomerContact
                        .Where(x => contactCodes.Contains(x.Code))
                        .ToListAsync();

                    // Cập nhật dữ liệu
                    contactsToUpdate.ForEach(entity =>
                    {
                        var updatedData = existingContacts.FirstOrDefault(c => c.Code == entity.Code);
                        if (updatedData != null)
                        {
                            entity.Customer_Code = updatedData.Customer_Code;
                            entity.Type = updatedData.Type;
                            entity.Value = updatedData.Value;
                            entity.IsActive = updatedData.IsActive;
                        }
                    });

                    // Dùng UpdateRange để cập nhật nhiều bản ghi cùng lúc
                    _dbContext.TblMdCustomerContact.UpdateRange(contactsToUpdate);
                }

                await _dbContext.SaveChangesAsync(); // Lưu dữ liệu
                await transaction.CommitAsync(); // Commit transaction

                var updatedContacts = _mapper.Map<List<CustomerContactDto>>(existingContacts);
                var insertedContacts = _mapper.Map<List<CustomerContactDto>>(newContacts);

                return insertedContacts.Concat(updatedContacts).ToList();
            }
            catch (Exception ex)
            {
                await transaction.RollbackAsync(); // Rollback nếu có lỗi
                Status = false;
                Exception = ex;
                return new List<CustomerContactDto>();
            }
        }




    }
}
