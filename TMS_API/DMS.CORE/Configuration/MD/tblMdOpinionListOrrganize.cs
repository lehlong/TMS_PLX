using DMS.CORE.Entities.AD;
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
    public class tblMdOpinionListOrrganize : IEntityTypeConfiguration<TblMdOpinionListOrganize>
    {
        public void Configure(EntityTypeBuilder<TblMdOpinionListOrganize> builder)
        {
            builder.HasKey(x => new { x.OpinionListCode, x.OrganizeId });
        }
    }
}
