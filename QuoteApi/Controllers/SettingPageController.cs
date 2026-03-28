using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Unicode;

namespace QuoteApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class SettingController : ControllerBase
    {
        private readonly IWebHostEnvironment _env;

        public SettingController(IWebHostEnvironment env)
        {
            _env = env;
        }

        [HttpGet("school/{schoolName}")]
        public IActionResult GetSchoolSetting(string schoolName)
        {
            var fileName = $"{schoolName.ToUpper()}.json";
            var path = Path.Combine(_env.WebRootPath, "files", fileName);

            if (!System.IO.File.Exists(path))
            {
                return NotFound($"找不到 {schoolName} 的設定檔");
            }

            var jsonString = System.IO.File.ReadAllText(path);
            return Ok(jsonString);
        }

        [HttpPost("school/{schoolName}")]
        public IActionResult SaveSchoolSetting(string schoolName, [FromBody] object settingData)
        {
            try
            {
                var fileName = $"{schoolName.ToUpper()}.json";
                var path = Path.Combine(_env.WebRootPath, "files", fileName);

                // 確保 files 資料夾存在
                var dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                var options = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    // 讓所有 Unicode 都直接輸出，不轉成 \uXXXX
                    Encoder = JavaScriptEncoder.Create(UnicodeRanges.All)
                };

                // 將 JSON 資料寫入檔案
                //var jsonString = JsonSerializer.Serialize(settingData, new JsonSerializerOptions { WriteIndented = true });
                var jsonString = JsonSerializer.Serialize(settingData, options);
                System.IO.File.WriteAllText(path, jsonString);

                return Ok(new { message = $"已儲存 {schoolName} 設定檔", fileName });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"儲存失敗: {ex.Message}");
            }
        }

    }
}
