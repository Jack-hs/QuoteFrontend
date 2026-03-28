using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;

[ApiController]
[Route("api/[controller]")]
public class ExchangeRateController : ControllerBase
{
    private readonly IMemoryCache _cache;

    public ExchangeRateController(IMemoryCache cache)
    {
        _cache = cache;
    }

    [HttpGet("latest")]
    public IActionResult GetLatestRates()
    {
        // 嘗試從快取中拿出剛才背景服務存進去的 "LatestRates"
        if (_cache.TryGetValue("LatestRates", out var rates))
        {
            return Ok(rates);
        }

        // 如果快取還沒準備好 (例如剛開機第一秒)，回傳一個提示
        return StatusCode(503, "匯率資料初始化中，請稍後再試。");
    }
}
