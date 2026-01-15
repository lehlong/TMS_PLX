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
using Hangfire.AspNetCore;
using DMS.BUSINESS.Services.AD;
using DMS.BUSINESS.Services.HUB;
using DMS.CORE;
using DMS.BUSINESS.Services.BackgroundHangfire;
using Microsoft.Extensions.FileProviders;

var config = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .AddEnvironmentVariables().Build();

var logger = LogManager.Setup()
    .LoadConfiguration(new NLogLoggingConfiguration(config.GetSection("NLog")))
    .GetCurrentClassLogger();

var builder = WebApplication.CreateBuilder(args);


// ==========================================
// 1) HANGFIRE CONFIG
// ==========================================
builder.Services.AddHangfire(options =>
{
    options.UseSqlServerStorage(builder.Configuration.GetConnectionString("HangfireConnection"));
});

// chạy server
builder.Services.AddHangfireServer();

// Đăng ký service job
builder.Services.AddScoped<BackgroundJobService>();

builder.Services.AddSingleton<IRecurringJobManager, RecurringJobManager>();


// ==========================================
// 2) API + JWT
// ==========================================
builder.Services.AddControllers();
builder.Services.AddDIServices(builder.Configuration);
// builder.Services.AddDIXHTDServices(builder.Configuration);

builder.Services.AddHttpContextAccessor();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddMvc();
builder.Services.AddSwaggerGen(options =>
{
    options.SwaggerDoc("V1", new OpenApiInfo
    {
        Version = "V1",
        Title = "WebAPI",
        Description = "<a href='/log' target='_blank'>Xem log file</a>"
    });

    options.AddSecurityDefinition("Bearer", new OpenApiSecurityScheme
    {
        Scheme = "Bearer",
        BearerFormat = "JWT",
        In = ParameterLocation.Header,
        Name = "Authorization",
        Description = "Bearer token",
        Type = SecuritySchemeType.Http
    });

    options.AddSecurityRequirement(new OpenApiSecurityRequirement {
        { new OpenApiSecurityScheme {
            Reference = new OpenApiReference {
                Id = "Bearer",
                Type = ReferenceType.SecurityScheme }
            }, new List<string>() }
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
})
.AddJwtBearer(options =>
{
    options.TokenValidationParameters = new TokenValidationParameters
    {
        ValidateIssuer = true,
        ValidateAudience = true,
        ValidateLifetime = true,
        ValidateIssuerSigningKey = true,
        ValidIssuer = config.GetValue<string>("JWT:Issuer"),
        ValidAudience = config.GetValue<string>("JWT:Audience"),
        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(config.GetValue<string>("JWT:Key"))),
        ClockSkew = TimeSpan.Zero
    };
});

builder.Services.Configure<FormOptions>(o =>
{
    o.ValueLengthLimit = int.MaxValue;
    o.MultipartBodyLengthLimit = int.MaxValue;
    o.MemoryBufferThreshold = int.MaxValue;
});

builder.Services.AddSignalR(options => options.EnableDetailedErrors = true);
builder.Services.AddMemoryCache();

builder.Services.AddCors(options => options.AddPolicy("CorsPolicy",
    policy =>
    {
        policy.WithOrigins("*")
            .AllowAnyMethod()
            .AllowAnyHeader();
    }));


var app = builder.Build();


// ==========================================
// 3) MIDDLEWARE
// ==========================================
app.UseHangfireDashboard("/hangfire");

app.UseSwagger();
app.UseSwaggerUI(options =>
{
    options.SwaggerEndpoint("/swagger/V1/swagger.json", "PROJECT WebAPI");
});

TransferObjectExtension.SetHttpContextAccessor(app.Services.GetRequiredService<IHttpContextAccessor>());
app.EnableRequestBodyRewind();

app.UseRouting();
app.UseCors("CorsPolicy");
app.UseAuthentication();
app.UseAuthorization();

app.UseMiddleware<ActionLoggingMiddleware>();

app.MapHub<SystemTraceServiceHub>("/SystemTrace");
app.MapHub<RefreshServiceHub>("/Refresh");

app.UseStaticFiles(new StaticFileOptions
{
    FileProvider = new PhysicalFileProvider(Path.Combine(Directory.GetCurrentDirectory(), "Uploads")),
    RequestPath = "/Uploads"
});

app.MapControllers();


// ==========================================
// 4) KHỞI TẠO RECURRING JOB ĐÚNG CÁCH
// ==========================================
var recurringJobs = app.Services.GetRequiredService<IRecurringJobManager>();

// Gọi đúng kiểu DI
RecurringJob.AddOrUpdate<BackgroundJobService>(
    "SendSMS",
    x => x.SendSMSAsync(),
    "*/30 * * * * *"
);

RecurringJob.AddOrUpdate<BackgroundJobService>(
    "SendMail",
    x => x.SendMailAsync(),
    "*/30 * * * * *"
);


// ==========================================
// 5) RUN APP
// ==========================================
app.Run();
