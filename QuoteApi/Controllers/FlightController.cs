using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json.Linq;

[ApiController]
[Route("api/[controller]")]
public class FlightController : ControllerBase
{
    private readonly HttpClient _http;
    private readonly IConfiguration _config;

    public FlightController(IHttpClientFactory httpFactory, IConfiguration config)
    {
        _http = httpFactory.CreateClient();
        _config = config;
    }

    [HttpGet("test")]
    public IActionResult Test()
    {
        var key = _config["SerpApi:Key"];
        return Ok(new
        {
            keyLoaded = !string.IsNullOrEmpty(key),
            keyPreview = key?.Length > 5 ? key[..5] + "..." : "空的！"
        });
    }

    [HttpGet("search")]
    public async Task<IActionResult> Search(
    [FromQuery] string from,
    [FromQuery] string to,
    [FromQuery] string outboundDate,
    [FromQuery] string? returnDate,
    [FromQuery] int adults = 1,
    [FromQuery] string tripType = "1")
    {
        // 👇 先檢查 Key 有沒有讀到
        var apiKey = _config["SerpApi:Key"];

        if (string.IsNullOrEmpty(apiKey))
            return BadRequest(new { message = "SerpAPI Key 未設定！請檢查 appsettings.json" });

        // 👇 印出 URL 方便 debug
        var url = $"https://serpapi.com/search.json" +
                  $"?engine=google_flights" +
                  $"&departure_id={from}" +
                  $"&arrival_id={to}" +
                  $"&outbound_date={outboundDate}" +
                  (returnDate != null ? $"&return_date={returnDate}" : "") +
                  $"&adults={adults}" +
                  $"&type={tripType}" +
                  $"&currency=TWD" +
                  $"&hl=zh-tw" +
                  $"&api_key={apiKey}";

        Console.WriteLine($"🔍 SerpAPI URL: {url}");  // 👈 看 Console 輸出

        try
        {
            var response = await _http.GetStringAsync(url);
            var json = JObject.Parse(response);

            // 👇 檢查 SerpAPI 有沒有回傳錯誤
            if (json["error"] != null)
                return BadRequest(new { message = json["error"]?.ToString() });

            var bestFlights = ParseFlights(json["best_flights"]);
            var otherFlights = ParseFlights(json["other_flights"]);
            var priceInsights = new
            {
                lowestPrice = json["price_insights"]?["lowest_price"]?.Value<int>() ?? 0,
                typicalRangeMin = json["price_insights"]?["typical_range"]?[0]?.Value<int>() ?? 0,
                typicalRangeMax = json["price_insights"]?["typical_range"]?[1]?.Value<int>() ?? 0,
            };

            return Ok(new { bestFlights, otherFlights, priceInsights });
        }
        catch (Exception ex)
        {
            Console.WriteLine($"💥 錯誤：{ex.Message}");
            return BadRequest(new { message = ex.Message });
        }
    }

    // ===== 解析航班資料 =====
    private List<object> ParseFlights(JToken? flightsToken)
    {
        var result = new List<object>();
        if (flightsToken == null) return result;

        foreach (var item in flightsToken)
        {
            // 取第一段航班資訊
            var firstFlight = item["flights"]?[0];
            var lastFlight = item["flights"]?.Last;

            var flightNumbers = item["flights"]
                ?.Select(f => f["flight_number"]?.ToString() ?? "")
                .Where(n => n != "")
                .ToList() ?? new List<string>();

            result.Add(new
            {
                price = item["price"]?.Value<int>() ?? 0,
                currency = "TWD",
                airline = firstFlight?["airline"]?.ToString() ?? "",
                airlineLogo = firstFlight?["airline_logo"]?.ToString() ?? "",
                totalDuration = item["total_duration"]?.Value<int>() ?? 0,
                stops = (item["flights"]?.Count() ?? 1) - 1,
                departureTime = firstFlight?["departure_airport"]?["time"]?.ToString() ?? "",
                arrivalTime = lastFlight?["arrival_airport"]?["time"]?.ToString() ?? "",
                flightNumbers,
                bookingToken = item["booking_token"]?.ToString() ?? "",
                type = item["type"]?.ToString() ?? "來回"
            });
        }

        return result;
    }
}
