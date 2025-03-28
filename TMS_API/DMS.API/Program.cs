using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.AspNetCore.Http.Features;
using Microsoft.IdentityModel.Tokens;
using DMS.BUSINESS;
using System.Text;
using Microsoft.OpenApi.Any;
using Microsoft.OpenApi.Models;
using DMS.API.AppCode.Extensions;
using NLog;
using NLog.Extensions.Logging;
using DMS.API.Middleware;
using Hangfire;
using Hangfire.Oracle.Core;
using DMS.BUSINESS.Services.AD;
using DMS.BUSINESS.Services.HUB;
using DMS.API.AppCode.Util;
using DMS.CORE;
using Common;
using Microsoft.Extensions.FileProviders;

var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddJsonFile($"appsettings.json", optional: true)
                .AddEnvironmentVariables().Build();
var logger = LogManager.Setup()
                       .LoadConfiguration(new NLogLoggingConfiguration(config.GetSection("NLog")))
                       .GetCurrentClassLogger();


var builder = WebApplication.CreateBuilder(args);
//builder.Services.AddHangfire(configuration =>
//            configuration.UseStorage(new OracleStorage(config.GetConnectionString("Connection"), new OracleStorageOptions())));

// Thêm dịch vụ Hangfire
//builder.Services.AddHangfireServer();

builder.Services.AddControllers();
builder.Services.AddDIServices(builder.Configuration);
//builder.Services.AddDIXHTDServices(builder.Configuration);
builder.Services.AddHttpContextAccessor();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddMvc();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("V1", new OpenApiInfo
    {
        Version = "V1",
        Title = "WebAPI",
        Description = "<a href='/log' target = '_blank'>Bấm vào đây để xem log file</a>",
    });
    options.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Scheme = "Bearer",
        BearerFormat = "JWT",
        In = ParameterLocation.Header,
        Name = "Authorization",
        Description = "Bearer Authentication with JWT Token",
        Type = SecuritySchemeType.Http
    });
    options.AddSecurityRequirement(new OpenApiSecurityRequirement {
        {
            new OpenApiSecurityScheme {
                Reference = new OpenApiReference {
                    Id = "Bearer",
                        Type = ReferenceType.SecurityScheme
                }
            },
            new List < string > ()
        }
    });
    options.MapType<TimeSpan>(() => new OpenApiSchema
    {
        Type = "string",
        Example = new OpenApiString("00:00:00")
    });
});

builder.Services.AddAuthentication(opt =>
{
    opt.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    opt.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
}).AddJwtBearer(options =>
{
    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateLifetime = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = config.GetSection("JWT:Issuer").Value,
        ValidAudience = config.GetSection("JWT:Audience").Value,
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(config.GetSection("JWT:Key").Value)),
        ClockSkew = TimeSpan.Zero
    };
});

builder.Services.Configure<FormOptions>(o =>
{
    o.ValueLengthLimit = int.MaxValue;
    o.MultipartBodyLengthLimit = int.MaxValue;
    o.MemoryBufferThreshold = int.MaxValue;
});

builder.Services.AddSignalR(options =>
{
    options.EnableDetailedErrors = true;
});

builder.Services.AddMemoryCache();

//builder.Services.AddCors(options => options.AddPolicy("CorsPolicy",
//        builder =>
//        {
//            builder.AllowAnyHeader()
//                    .AllowAnyMethod()
//                    .AllowCredentials()
//                    .SetIsOriginAllowed((host) => true);
//        }));
builder.Services.AddCors(options => options.AddPolicy("CorsPolicy",
        builder =>
        {
            builder.WithOrigins("*")
           .AllowAnyMethod()
           .AllowAnyHeader();
        }));

var app = builder.Build();

//if (!app.Environment.IsDevelopment())
//{
//    //app.UseHangfireDashboard();
//   // using var scope = app.Services.CreateScope();
//   // using var server = new BackgroundJobServer();
//   // await scope.ServiceProvider.GetRequiredService<ISystemTraceService>().StartService();
//   // var dbContext = scope.ServiceProvider.GetRequiredService<AppDbContext>();
//   // var lstMessage = dbContext.TblAdMessage.ToList();
//    foreach (var message in lstMessage)
//    {
//        MessageUtil.AddToCache(new MessageObject()
//        {
//            Code = message.Code,
//            Language = message.Lang,
//            Message = message.Value
//        });
//    }
//}

// Configure the HTTP request pipeline.
//if (app.Environment.IsDevelopment())
//{
//    app.UseSwagger();
//    app.UseSwaggerUI(options =>
//    {
//        options.SwaggerEndpoint("/swagger/V1/swagger.json", "PROJECT WebAPI");
//    });
//}
app.UseSwagger();
app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint("/swagger/V1/swagger.json", "PROJECT WebAPI");
});

TransferObjectExtension.SetHttpContextAccessor(app.Services.GetRequiredService<IHttpContextAccessor>());
app.EnableRequestBodyRewind();

//app.UseHttpsRedirection();


app.UseRouting();
app.UseCors("CorsPolicy");

app.UseAuthentication();

app.UseAuthorization();


app.UseMiddleware<ActionLoggingMiddleware>();
app.MapHub<SystemTraceServiceHub>("/SystemTrace");
app.MapHub<RefreshServiceHub>("/Refresh");

app.UseStaticFiles();

app.MapControllers();
app.Run();
