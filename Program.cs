using GREETApi.Services;
using GREETDataAccess;
using Microsoft.EntityFrameworkCore;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

// CORS configuration to allow Blazor Server app
builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(
        builder =>
        {
            builder.WithOrigins("https://localhost:5300") // Replace with Blazor Server app URL if different
                   .AllowAnyHeader()
                   .AllowAnyMethod();
        });
});

// Configure DbContext with PostgreSQL
builder.Services.AddDbContext<PostGreSQLDbContext>(options =>
    options.UseNpgsql(builder.Configuration.GetConnectionString("DefaultConnection")));

// Register ExcelService
builder.Services.AddSingleton(provider =>
    new ExcelService());

var app = builder.Build();

// Apply migrations at runtime
using (var scope = app.Services.CreateScope())
{
    var dbContext = scope.ServiceProvider.GetRequiredService<PostGreSQLDbContext>();
    dbContext.Database.Migrate();
}

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();
app.UseCors(); // Enable CORS

app.UseAuthorization();

app.MapControllers();

app.Run();
