var builder = WebApplication.CreateBuilder(args);
builder.Services.AddHttpClient();  // 👈 加這行
builder.Services.AddControllers();
builder.Services.AddCors(options => 
{
    options.AddPolicy("AllowAll", 
        policy => policy.AllowAnyOrigin()
                       .AllowAnyMethod()
                       .AllowAnyHeader());
});

var app = builder.Build();

app.UseCors("AllowAll");
app.UseStaticFiles();  // 讓 wwwroot/files 可讀

app.MapControllers();

app.Run();
