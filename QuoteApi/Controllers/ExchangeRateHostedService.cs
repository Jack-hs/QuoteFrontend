using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Threading;
using System.Threading.Tasks;

public class ExchangeRateHostedService : BackgroundService
{
    private readonly ILogger<ExchangeRateHostedService> _logger;
    private readonly IServiceProvider _serviceProvider;
    private readonly IMemoryCache _cache;

    public ExchangeRateHostedService(
        ILogger<ExchangeRateHostedService> logger,
        IServiceProvider serviceProvider,
        IMemoryCache cache)
    {
        _logger = logger;
        _serviceProvider = serviceProvider;
        _cache = cache;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        _logger.LogInformation("匯率抓取背景服務已啟動。");

        // 設定間隔時間：1 小時 (你可以改成 TimeSpan.FromMinutes(1) 來測試)
        using PeriodicTimer timer = new PeriodicTimer(TimeSpan.FromHours(1));

        try
        {
            // 服務啟動時先抓第一次，免得 API 被呼叫時還沒有快取資料
            await FetchAndCacheRatesAsync();

            // 進入迴圈：每次等待 timer 的週期時間到，就會往下執行
            while (await timer.WaitForNextTickAsync(stoppingToken))
            {
                await FetchAndCacheRatesAsync();
            }
        }
        catch (OperationCanceledException)
        {
            _logger.LogInformation("匯率抓取背景服務已停止。");
        }
    }

    private async Task FetchAndCacheRatesAsync()
    {
        try
        {
            _logger.LogInformation($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] 開始抓取最新匯率...");

            // 因為 BackgroundService 是 Singleton，但 HttpClient 通常是 Transient 或 Scoped
            // 所以我們用 CreateScope() 來解析剛才寫好的 ExchangeRateService
            using var scope = _serviceProvider.CreateScope();
            var exchangeService = scope.ServiceProvider.GetRequiredService<ExchangeRateService>();

            // 抓取資料
            var fubonUsd = await exchangeService.GetFubonUsdRateAsync();
            var botPhp = await exchangeService.GetBotPhpRateAsync();

            //// 將資料寫入 MemoryCache，快取有效時間我們設為 65 分鐘 (比抓取週期長一點點比較保險)
            //_cache.Set("LatestRates", new
            //{
            //    LastUpdated = DateTime.Now,
            //    USD = new { Bank = "富邦銀行", BuyRate = fubonUsd.BuySpot, SellRate = fubonUsd.SellSpot },
            //    PHP = new { Bank = "台灣銀行", BuyRate = botPhp.BuyCash, SellRate = botPhp.SellCash }
            //}, TimeSpan.FromMinutes(65));
            decimal.TryParse(fubonUsd.BuySpot, out decimal BuySpot);
            decimal.TryParse(fubonUsd.SellSpot, out decimal SellSpot);
            decimal.TryParse(botPhp.BuyCash, out decimal BuyCash);
            decimal.TryParse(botPhp.SellCash, out decimal SellCash);
            _cache.Set("LatestRates", new ExchangeRates
            {
                LastUpdated = DateTime.Now,
                USD = new CurrencyRate
                {
                    Bank = "富邦銀行",
                    BuyRate = BuySpot,
                    SellRate = SellSpot
                },
                PHP = new CurrencyRate
                {
                    Bank = "台灣銀行",
                    BuyRate = BuyCash,
                    SellRate = SellCash
                }
            }, TimeSpan.FromMinutes(65));

            _logger.LogInformation("最新匯率已成功抓取並更新至快取。");
        }
        catch (Exception ex)
        {
            _logger.LogError(ex, "抓取匯率時發生錯誤！");
        }
    }
}
