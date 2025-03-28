using DMS.CORE.Entities.MD;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Metadata.Builders;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DMS.CORE.Configuration.MD
{
    public class TblMdAuditTemplateListTablesDataConfig : IEntityTypeConfiguration<TblMdAuditTemplateListTablesData>
    {
        public void Configure(EntityTypeBuilder<TblMdAuditTemplateListTablesData> builder)
        {
            builder.HasKey(x => new { x.TemplateDataCode, x.AuditListTablesCode });
        }
    }
}
