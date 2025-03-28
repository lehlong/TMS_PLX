﻿using Microsoft.Extensions.DependencyInjection;
using DMS.CORE;
using DMS.BUSINESS.Services.AD;
using System.Reflection;
using Microsoft.Extensions.Configuration;
using AutoMapper.Extensions.ExpressionMapping;
using Microsoft.EntityFrameworkCore;
using DMS.BUSINESS.Common;

namespace DMS.BUSINESS
{
    public static class ServiceExtension
    {
        public static IServiceCollection AddDIServices(this IServiceCollection services, IConfiguration configuration)
        {
            services.AddAutoMapper(cfg => { cfg.AddExpressionMapping(); }, typeof(MappingProfile).Assembly);
            // Add Entity Framework
            //services.AddDbContext<AppDbContext>(options => options.UseOracle(configuration.GetConnectionString("Connection"), b =>
            //         b.UseOracleSQLCompatibility(OracleSQLCompatibility.DatabaseVersion19)));
            services.AddDbContext<AppDbContext>(options => options.UseSqlServer(configuration.GetConnectionString("Connection")));
            //Add all service
            var allProviderTypes = Assembly.GetAssembly(typeof(IAccountService))
             .GetTypes().Where(t => t.Namespace != null).ToList();
            foreach (var intfc in allProviderTypes.Where(t => t.IsInterface))
            {
                var impl = allProviderTypes.FirstOrDefault(c => c.IsClass && !c.IsAbstract && intfc.Name[1..] == c.Name);
                if (impl != null) services.AddScoped(intfc, impl);
            }

            return services;
        }
    }
}
