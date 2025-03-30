using DMS.CORE.Common;
using DMS.CORE.Entities.AD;
using DMS.CORE.Entities.BU;
using DMS.CORE.Entities.IN;
using DMS.CORE.Entities.MD;

using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;

namespace DMS.CORE
{
    public class AppDbContext(DbContextOptions<AppDbContext> options, IHttpContextAccessor httpContextAccessor) : DbContext(options)
    {
        public IHttpContextAccessor _httpContextAccessor = httpContextAccessor;

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.ApplyAllConfigurations();
            foreach (var type in modelBuilder.Model.GetEntityTypes())
            {
                if (typeof(ISoftDeleteEntity).IsAssignableFrom(type.ClrType))
                    modelBuilder.SetSoftDeleteFilter(type.ClrType);
            }

            modelBuilder.HasSequence<int>("ORDER_SEQUENCE")
                    .StartsAt(1)
                    .IncrementsBy(1);

            base.OnModelCreating(modelBuilder);
        }

        public Func<DateTime> TimestampProvider { get; set; } = ()
            => DateTime.Now;

        public override int SaveChanges()
        {
            TrackChanges();
            return base.SaveChanges();
        }

        public override async Task<int> SaveChangesAsync(CancellationToken cancellationToken = default)
        {
            TrackChanges();
            return await base.SaveChangesAsync(cancellationToken);
        }
        public string GetUserRequest()
        {
            var tokens = _httpContextAccessor?.HttpContext?.Request?.Headers.Authorization.ToString()?.Split(" ")?.ToList();
            string? user = null;
            if (tokens != null)
            {
                var token = tokens.FirstOrDefault(x => x != "Bearer");
                if (!string.IsNullOrWhiteSpace(token) && token != "null")
                {
                    JwtSecurityTokenHandler tokenHandler = new();
                    JwtSecurityToken securityToken = tokenHandler.ReadToken(token) as JwtSecurityToken;
                    var claim = securityToken.Claims;
                    var result = claim.FirstOrDefault(x => x.Type == ClaimTypes.Name);
                    user = result?.Value;
                }
            }
            return user;
        }

        private void TrackChanges()
        {
            var user = GetUserRequest();

            foreach (var entry in ChangeTracker.Entries().Where(e => e.State == EntityState.Added || e.State == EntityState.Modified))
            {
                if (entry.Entity is IBaseEntity auditable)
                {
                    if (entry.State == EntityState.Added)
                    {
                        auditable.CreateBy = user;
                        auditable.CreateDate = TimestampProvider();
                    }
                    else
                    {
                        Entry(auditable).Property(x => x.CreateBy).IsModified = false;
                        Entry(auditable).Property(x => x.CreateDate).IsModified = false;
                        auditable.UpdateBy = user;
                        auditable.UpdateDate = TimestampProvider();
                    }
                }
            }

            foreach (var entry in ChangeTracker.Entries().Where(e => e.State == EntityState.Deleted))
            {
                if (entry.Entity is ISoftDeleteEntity deletedEntity)
                {
                    entry.State = EntityState.Unchanged;
                    deletedEntity.IsDeleted = true;
                    deletedEntity.DeleteBy = user;
                    deletedEntity.DeleteDate = TimestampProvider();
                }
            }
        }

        public async Task<int> GetNextSequenceValue(string sequence)
        {
            using var command = Database.GetDbConnection().CreateCommand();
            command.CommandText = $"SELECT {sequence}.NEXTVAL FROM DUAL";
            await Database.OpenConnectionAsync();
            using var result = await command.ExecuteReaderAsync();
            await result.ReadAsync();
            return result.GetInt32(0);
        }

        #region System Manage
        public DbSet<TblAdConfigTemplate> TblAdConfigTemplate { get; set; }

        public DbSet<TblAdAccount> TblAdAccount { get; set; }
        public DbSet<TblAdAccountGroup> TblAdAccountGroup { get; set; }
        public DbSet<TblAdMenu> TblAdMenu { get; set; }
        public DbSet<TblAdRight> TblAdRight { get; set; }
        public DbSet<TblAdMessage> TblAdMessage { get; set; }
        public DbSet<TblAdAccountGroupRight> TblAdAccountGroupRight { get; set; }
        public DbSet<TblAdAccountRefreshToken> TblAdAccountRefreshToken { get; set; }
        public DbSet<TblAdAppVersion> TblAdAppVersion { get; set; }
        public DbSet<TblAdAccount_AccountGroup> TblAdAccount_AccountGroup { get; set; }
        public DbSet<TblActionLog> TblActionLogs {get; set;}
        public DbSet<TblAdSystemTrace> TblAdSystemTrace { get; set; }
        public DbSet<tblAdOrganize> tblAdOrganize { get; set; }
        #endregion

        #region Master Data
        public DbSet<TblMdDeliveryGroup> TblMdDeliveryGroup { get; set; }
        public DbSet<TblMdTermOfPayment> TblMdTermOfPayment { get; set; }
        public DbSet<TblMdCustomerDb> TblMdCustomerDb { get; set; }
        public DbSet<TblMdCustomerPt> TblMdCustomerPt { get; set; }
        public DbSet<TblMdCustomerFob> TblMdCustomerFob { get; set; }
        public DbSet<TblMdCustomerTnpp> TblMdCustomerTnpp { get; set; }
        public DbSet<TblMdCustomerBbdo> TblMdCustomerBbdo { get; set; }




        public DbSet<TblMdGiaGiaoTapDoanList> TblMdGiaGiaoTapDoanList { get; set; }
        public DbSet<TblMdTypeOfGoods> TblMdTypeOfGoods { get; set; }
        public DbSet<TblMdGoods> TblMdGoods { get; set; }
        public DbSet<TblMdWarehouse> TblMdWarehouse { get; set; }
        public DbSet<TblMdLaiGopDieuTiet> TblMdLaiGopDieuTiet { get; set; }
        public DbSet<TblMdCustomer> TblMdCustomer { get; set; }
        public DbSet<TblMdCompetitor> TblMdCompetitor { get; set; }
        public DbSet<TblMdMapPointCustomerGoods> TblMdMapPointCustomerGoods { get; set; }
        public DbSet<TblMdMarketCompetitor> TblMdMarketCompetitor { get; set; }

        public DbSet<TblMdDeliveryPoint> TblMdDeliveryPoint { get; set; }
        public DbSet<TblMdGiaGiaoTapDoan> TblMdGiaGiaoTapDoan { get; set; }
        public DbSet<TblMdCustomerType> TblMdCustomerType { get; set; }

        public DbSet<TblMdArea> TblMdArea { get; set; }
      
        public DbSet<TblMdCurrency> tblMdCurrency { get; set; }
        public DbSet<TblMdUnit> tblMdUnit { get; set; }

        public DbSet<TblMdLocal> tblMdLocal { get; set; }

        public DbSet<TblMdPeriodTime> tblMdPeriodTime { get; set; }
       
        public DbSet<TblMdOpinionType> tblMdOpinionType { get; set; }
        public DbSet<TblMdTemplateListTables> tblMdTemplateListTables { get; set; }
        public DbSet<TblMdTemplateListTablesGroups> tblMdTemplateListTablesGroups { get; set; }
        public DbSet<TblMdTemplateListTablesData> tblMdTemplateListTablesData { get; set; }
        public DbSet<TblMdAuditPeriod> tblMdAuditPeriods { get; set; }
        public DbSet<TblMdMgListTablesGroups> tblMdMgListTablesGroup { get; set; }
        public DbSet<TblMdMgListTables> tblMdMgListTables { get; set; }
        public DbSet<TblMdMgOpinionList> tblMdMgOpinionList { get; set; }
        public DbSet<TblMdAccountType> tblMdAccountType { get; set; }
        public DbSet<TblMdListAudit> tblMdListAudit { get; set; }
        public DbSet<tblMdListAuditHistory> tblMdListAuditHistory { get; set; }
        public DbSet<TblMdAuditTemplateHistory> tblMdAuditTemplateHistory { get; set; }
        public DbSet<TblMdOpinionListOrganize> tblMdOpinionListOrganize { get; set; }
        public DbSet<TblMdAuditPeriodListTables> tblMdAuditPeriodListTables { get; set; }
        public DbSet<TblMdAuditTemplateListTablesData> tblMdAuditTemplateListTablesData { get; set; }
        public DbSet<TblMdSalesMethod> TblMdSalesMethod { get; set; }
        public DbSet<TblMdRetailPrice> TblMdRetailPrice { get; set; }
        public DbSet<TblMdRetailPriceList> TblMdRetailPriceList { get; set; }
        public DbSet<TblMdMarket> TblMdMarket { get; set; }
        public DbSet<TblMdCuocVanChuyen> TblMdCuocVanChuyen { get; set; }
        public DbSet<TblMdCuocVanChuyenList> TblMdCuocVanChuyenList { get; set; }

        public DbSet<TblMdSigner> TblMdSigner { get; set; }
        public DbSet<TblMdCustomerContact> TblMdCustomerContact { get; set; }

        public DbSet<TblMdTemplateReport> tblMdTemplateReport { get; set; }
        public DbSet<TblMdTemplateReportElement> tblMdTemplateReportElement { get; set; }
        public DbSet<TblMdTemplateReportMapping> tblMdTemplateReportMapping { get; set; }




        #endregion

        #region Input
        public DbSet<TblInNguoiKyTen> TblInNguoiKyTen { get; set; }
        public DbSet<TblInVinhCuaLo> TblInVinhCuaLo { get; set; }
        public DbSet<TblInHeSoMatHang> TblInHeSoMatHang { get; set; }
        public DbSet<TblInDiscountCompetitor> TblInDiscountCompetitor { get; set; }
        public DbSet<TbLInCustomerFob> TbLInCustomerFob { get; set; }
        public DbSet<TblInDiscountCompany> TblInDiscountCompany { get; set; }

        #endregion

        #region Common
        public DbSet<TblCmAttachment> TblBuAttachment { get; set; }
        public DbSet<TblCmModuleAttachment> TblBuModuleAttachment { get; set; }
        public DbSet<TblCmComment> TblCmComment { get; set; }
        public DbSet<TblCmModuleComment> TblCmModuleComment { get; set; }
        public DbSet<TblNotifyEmail> TblCmNotifiEmail{ get; set; }
        public DbSet<TblNotifySms> TblCmNotifySms { get; set; }

        #endregion

        #region Business
        public DbSet<tblBuOpinionList> TblBuOpinionLists { get; set; }
        public DbSet<TblBuHistoryDownload> TblBuHistoryDownload { get; set; }
        public DbSet<TblBuHistoryAction> TblBuHistoryAction { get; set; }
        public DbSet<TblBuListTables> TblBuListTables { get; set; }
        public DbSet<TblBuCalculateResultList> TblBuCalculateResultList { get; set; }
        public DbSet<tblBuOpinionDetail> tblBuOpinionDetail { get; set; }
        public DbSet<tblBuOpinionDetailHistory> tblBuOpinionDetailHistory { get; set; }
        public DbSet<tblBuPendingOpinionMapping> tblBuPendingOpinionMappings { get; set; }
        public DbSet<TblBuDiscountInformationList> TblBuDiscountInformationList { get; set; }





        public DbSet<TblBuCalculateDiscount> TblBuCalculateDiscount { get; set; }
        public DbSet<TblBuInputPrice> TblBuInputPrice { get; set; }
        public DbSet<TblBuInputMarket> TblBuInputMarket { get; set; }
        public DbSet<TblBuInputCustomerDb> TblBuInputCustomerDb { get; set; }
        public DbSet<TblBuInputCustomerFob> TblBuInputCustomerFob { get; set; }
        public DbSet<TblBuInputCustomerTnpp> TblBuInputCustomerTnpp { get; set; }
        public DbSet<TblBuInputCustomerPt> TblBuInputCustomerPt { get; set; }
        public DbSet<TblBuInputCustomerBbdo> TblBuInputCustomerBbdo { get; set; }
        #endregion
    }
}
