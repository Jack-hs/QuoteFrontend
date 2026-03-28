using QuoteApi.Models;

var builder = WebApplication.CreateBuilder(args);
// 1. 註冊 MemoryCache
builder.Services.AddMemoryCache();

// 2. 註冊 HttpClient 與爬蟲服務
builder.Services.AddHttpClient<ExchangeRateService>();

// 3. 註冊背景 Timer 服務
builder.Services.AddHostedService<ExchangeRateHostedService>();

// Add services to the container.
builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
// 在 builder.Build() 之前設定 EPPlus 授權
// 對於 EPPlus 8.x 及更高版本，請使用以下方法設定授權
OfficeOpenXml.ExcelPackage.License.SetNonCommercialPersonal("Your Full Name");

builder.Services.AddHttpClient();  // 👈 加這行
builder.Services.AddControllers();
builder.Services.AddCors(options => 
{
    options.AddPolicy("AllowAll", 
        policy => policy.WithOrigins(
            "https://jack-hs.github.io/QuoteFrontend", //你的 GitHub Pages
            "https://carma-proemployer-elnora.ngrok-free.dev") // 你的 ngrok（前端用不到，但後備）)
                       .AllowAnyMethod()
                       .AllowAnyHeader());
});

var app = builder.Build();

app.UseCors("AllowAll");
app.UseStaticFiles();  // 讓 wwwroot/files 可讀

// ==========================================
// 新增：在應用程式啟動時初始化全域設定與 ZoneData
// ==========================================
var env = app.Environment;
ZoneData.StartupPath = env.WebRootPath;
LoadSchoolIniAtStartup(env.WebRootPath);

app.MapControllers();
app.Run();

// 將原本寫在 Controller 裡的邏輯搬移到這裡
void LoadSchoolIniAtStartup(string webRootPath)
{
    var iniPath = Path.Combine(webRootPath, "files", "SchoolList.ini");
    var schoolIni = new QuoteApi.Models.IniFile(iniPath); // 請確保引用正確的 namespace

    ZoneData.QuotationTermsDict.Clear();
    int QuotationTerms = schoolIni.GetKeys("QuotationTerms").Count();

    for (int i = 1; i <= QuotationTerms; i++)
    {
        string Terms = schoolIni.IniReadUTF8("QuotationTerms", "Term" + i);
        if (!ZoneData.QuotationTermsDict.ContainsKey(Terms))
        {
            ZoneData.QuotationTermsDict.Add(Terms, new List<Tuple<int, int>>());
        }

        int ranges = schoolIni.GetKeys("Term" + i).Count();
        if (ranges > 0)
        {
            for (int j = 1; j <= ranges; j++)
            {
                string context = schoolIni.IniReadUTF8("Term" + i, "range" + j);
                ZoneData.QuotationTermsDict[Terms].Add(new Tuple<int, int>(Terms.IndexOf(context) + 1, context.Count()));
            }
        }
    }
    Console.WriteLine("QuotationTerms Config Init Completed at Startup");
}