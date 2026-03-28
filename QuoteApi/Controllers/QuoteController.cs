using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using QuoteApi.Models;  // IniFile 命名空間
using System.ComponentModel;
using System.Dynamic;  // 新增這行
using System.IO;
using System.Text.Json;
using static QuoteApi.Models.TuitionData;


namespace QuoteApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class QuoteController : ControllerBase
    {
        
        //public IniFile _schoolIni;
        private readonly IWebHostEnvironment _env;  // 注入！
        private readonly IniFile _schoolIni;
        private readonly OperationExcel operationExcel;
        private readonly TuitionFeeCalculate tuitionFeeCalculate;
        AcademyData academyData;
        //OperationExcel operationExcel;
        //TuitionFeeCalculate tuitionFeeCalculate;


        // 加建構子
        public QuoteController(IWebHostEnvironment env)
        {
            _env = env;
            ZoneData.StartupPath = env.WebRootPath;
            _schoolIni = new IniFile(Path.Combine(ZoneData.StartupPath, "files", "SchoolList.ini"));
            //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            operationExcel = new OperationExcel();
            tuitionFeeCalculate = new TuitionFeeCalculate();
            //LoadschoolIni();
        }

        [HttpPost("upload-excel")]
        public IActionResult UploadExcel([FromForm] IFormFile file, [FromForm] string schoolName, [FromForm] string fileType)
        {
            if (file == null || file.Length == 0) return BadRequest("請選擇要上傳的檔案");
            if (string.IsNullOrWhiteSpace(schoolName)) return BadRequest("請提供學校名稱");
            if (string.IsNullOrWhiteSpace(fileType)) return BadRequest("請選擇檔案類型 (Tuition 或 LocalFee)");

            try
            {
                // 確保 files 資料夾存在
                var filesDir = Path.Combine(_env.WebRootPath, "files");
                if (!Directory.Exists(filesDir)) Directory.CreateDirectory(filesDir);

                // 決定輸出的 JSON 檔名
                string suffix = fileType.Equals("tuition", StringComparison.OrdinalIgnoreCase) ? "TuitionData" : "LocalFeeData";
                string outputName = $"{schoolName.ToUpper()}_{suffix}.json";
                string outputPath = Path.Combine(filesDir, outputName);

                // 開啟上傳檔案的資料流
                using (var stream = file.OpenReadStream())
                {
                    if (fileType.Equals("tuition", StringComparison.OrdinalIgnoreCase))
                    {
                        var baseData = operationExcel.ParseTUITIONExcelToJsonFromStream(stream, outputPath);

                        // 如果想要確保兩份檔案都上傳後才產生系統檔，可以做個簡單判斷：
                        string localFeePath = Path.Combine(filesDir, $"{schoolName.ToUpper()}_LocalFeeData.json");
                        string tuitionPath = Path.Combine(filesDir, $"{schoolName.ToUpper()}_TuitionData.json");
                        // 當發現兩份檔案都存在時，就進行合併產生
                        if (System.IO.File.Exists(localFeePath) && System.IO.File.Exists(tuitionPath))
                        {
                            operationExcel.GenerateSystemJsonFromTemplate(schoolName, baseData);
                        }
                    }
                    else
                    {
                        operationExcel.ParseLocalFeeToJsonFromStream(stream, outputPath);
                    }
                }

                return Ok(new { message = "轉換並儲存成功！", fileName = outputName });
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"解析失敗: {ex.Message}");
            }
        }


        private void LoadschoolIni()
        {
            //修改報價須知ini
            ZoneData.QuotationTermsDict.Clear();
            int QuotationTerms = _schoolIni.GetKeys("QuotationTerms").Count();
            for (int i = 1; i <= QuotationTerms; i++)
            {
                //if (i == 0) continue;
                string Terms = _schoolIni.IniReadUTF8("QuotationTerms", "Term" + i);
                if (!ZoneData.QuotationTermsDict.ContainsKey(Terms)) ZoneData.QuotationTermsDict.Add(Terms, new List<Tuple<int, int>>());

                int ranges = _schoolIni.GetKeys("Term" + i).Count();
                if (ranges > 0) //有要範圍要變換紅色字強調
                {
                    for (int j = 1; j <= ranges; j++)
                    {
                        string context = _schoolIni.IniReadUTF8("Term" + i, "range" + j);
                        ZoneData.QuotationTermsDict[Terms].Add(new Tuple<int, int>(Terms.IndexOf(context) + 1, context.Count()));
                    }

                }
            }
            Console.WriteLine("QuotationTerms Config Finish");
        }

        private readonly Dictionary<string, string> _schoolTuitionFiles = new(StringComparer.OrdinalIgnoreCase)
        {
            { "PHILINTER", "PHILINTER-TUITION-updated-on-November-25-2025-3.xlsx" },
            { "EV", "EV-TUITION.xlsx" },
            { "PINES", "PINES-TUITION.xlsx" },
            { "JIC", "JIC-TUITION.xlsx" }
        };

        /// <summary>
        /// 取得該學校 Excel 內的所有頁籤名稱
        /// </summary>
        [HttpGet("school-sheets/{schoolName}")]
        public IActionResult GetSchoolSheets(string schoolName)
        {
            academyData = new AcademyData();
            try
            {
                // 1. 檢查是否有對應的檔案
                if (!_schoolTuitionFiles.TryGetValue(schoolName, out string fileName))
                {
                    return NotFound($"找不到學校 {schoolName} 的設定檔。");
                }

                //// 2. 組合完整路徑
                //var filePath = Path.Combine(_env.WebRootPath, "files", fileName);

                //if (!System.IO.File.Exists(filePath))
                //{
                //    return NotFound($"伺服器上找不到檔案: {fileName}");
                //}

                //// 3. 讀取 Excel 檔案
                //using var package = new ExcelPackage(new FileInfo(filePath));

                //// 4. 抓取所有頁籤(Worksheet)的名稱
                //var sheetNames = package.Workbook.Worksheets
                //                        .Select(ws => ws.Name)
                //                        .ToList();
                string[] sheetNames = null;
                return Ok(sheetNames);
            }
            catch (Exception ex)
            {
                return StatusCode(500, $"讀取 Excel 失敗: {ex.Message}");
            }
        }

        [HttpGet("school-list")]
        public IActionResult GetSchoolList()
        {
            //operationExcel.ParseLocalFeeToJson();
            //operationExcel.ParseTUITIONExcelToJson();
            try
            {
                var schools = _schoolIni.GetKeys("SchoolName")
                        .Select(k => _schoolIni.ReadValue("SchoolName", k))
                        .ToArray();
                return Ok(schools);
            }
            catch (Exception ex)
            {
                return Ok(new[] { "Read Faile" });
            }
        }

        // 接收前端傳來的 schoolName
        [HttpGet("school-details")]
        public IActionResult GetSchoolDetails([FromQuery] string schoolName)
        {
            if (string.IsNullOrWhiteSpace(schoolName)) return BadRequest("請提供學校名稱");

            // 1. 根據學校名稱去抓取對應的 tuition JSON 檔案
            // 假設檔案命名規則是：學校名_tuition.json，或者你目前統一用 tuition_data.json
            var fileName = $"{schoolName.ToUpper()}_TuitionData.json"; // 之後可以改成 $"{schoolName}_tuition.json"
            var path = Path.Combine(_env.WebRootPath, "files", fileName);

            if (!System.IO.File.Exists(path)) return NotFound($"找不到學校 {schoolName} 的學費資料");

            var jsonString = System.IO.File.ReadAllText(path);

            // 這裡使用動態解析 (JsonDocument) 是最簡單的方式，因為可以自動展開動態的 Key
            using JsonDocument doc = JsonDocument.Parse(jsonString);
            var root = doc.RootElement;

            // 用 Hashset 來避免重複的項目
            var courseNames = new HashSet<string>();
            var roomTypes = new HashSet<string>();
            var PlaceofStays = new HashSet<string>();

            // 2. 遍歷 tuition_data.json 的結構提取資料
            // 結構：大分類(ESL) -> 課程(INTENSIVE ESL) -> 房型(Single)
            foreach (var category in root.EnumerateObject())
            {
                foreach (var course in category.Value.EnumerateObject())
                {
                    // 加入課程名稱 (例如 "IELTS INTENSIVE")
                    courseNames.Add(course.Name);

                    // 進入課程，抓取房型 (例如 "Single", "Double")
                    foreach (var room in course.Value.EnumerateObject())
                    {
                        roomTypes.Add(room.Name);
                    }
                }
            }

            //找LocalFee
            fileName = $"{schoolName.ToUpper()}_LocalFeeData.json"; // 之後可以改成 $"{schoolName}_tuition.json"
            path = Path.Combine(_env.WebRootPath, "files", fileName);

            if (!System.IO.File.Exists(path)) return NotFound($"找不到學校 {schoolName} 的學雜費資料");
            jsonString = System.IO.File.ReadAllText(path);

            // 這裡使用動態解析 (JsonDocument) 是最簡單的方式，因為可以自動展開動態的 Key
            using JsonDocument doc2 = JsonDocument.Parse(jsonString);
            var root2 = doc2.RootElement;

            // 2. 遍歷 tuition_data.json 的結構提取資料
            // 結構：大分類(DORMITORY) -> 課程(Special Study Permit SSP) -> 房型(Single)
            foreach (var category in root2.EnumerateObject())
            {
                PlaceofStays.Add(category.Name);
            }
            

            // 3. 回傳給前端
            return Ok(new
            {
                courses = courseNames.ToList(),
                rooms = roomTypes.ToList(),
                placeofstays = PlaceofStays.ToList()
            });
        }

        [HttpGet("schools")]
        public IActionResult GetSchools()
        {
            return Ok(new[] { "PHILINTER", "EV", "PINES", "JIC" });
        }

        

        [HttpPost("calculate")]
        public IActionResult Calculate([FromBody] QuoteRequest request)
        {
            List<CourseFeeItem> courseFeeItems = new List<CourseFeeItem>();
            List<CourseFeeItem> localFeeList = new List<CourseFeeItem>();
            string schoolName = request.School.ToUpper();
            //讀取目前新的ini 資料
            operationExcel.ParseSchoolJson($"{schoolName}"); //主要讓 ZoneData.appSettings 賦予值
            

            // 1. 讀取 tuition_data.json
            var path = Path.Combine(_env.WebRootPath, "files", $"{schoolName}_TuitionData.json");
            var jsonString = System.IO.File.ReadAllText(path);
            using JsonDocument doc = JsonDocument.Parse(jsonString);
            var root = doc.RootElement;

            decimal tuitionPrice = 0;
            string targetWeek = $"{request.Weeks}W";

            //// 2. 尋找對應的學費與住宿費 (合併在 tuition_data.json 中)
            //foreach (var category in root.EnumerateObject())
            //{
            //    if (category.Value.TryGetProperty(request.Course, out var courseNode))
            //    {
            //        if (courseNode.TryGetProperty(request.RoomType, out var roomNode))
            //        {
            //            if (roomNode.TryGetProperty(targetWeek, out var priceElement))
            //            {
            //                tuitionPrice = priceElement.GetDecimal();
            //            }
            //        }
            //    }
            //}
            // 把前端傳來的字串先去頭去尾，並轉小寫，方便等一下比對
            string reqCourse = request.Course?.Trim().ToLower() ?? "";
            string reqRoomType = request.RoomType?.Trim().ToLower() ?? "";

            // 2. 尋找對應的學費與住宿費 (使用高容錯遍歷法)
            foreach (var category in root.EnumerateObject())
            {
                foreach (var course in category.Value.EnumerateObject())
                {
                    // 把 JSON 裡的課程名稱也去頭去尾轉小寫
                    string jsonCourseName = course.Name.Trim().ToLower();

                    // 如果名稱吻合 (例如都變成 "ielts intensive")
                    if (jsonCourseName == reqCourse)
                    {
                        foreach (var room in course.Value.EnumerateObject())
                        {
                            string jsonRoomName = room.Name.Trim().ToLower();

                            // 如果房型名稱吻合 (例如都變成 "single")
                            if (jsonRoomName == reqRoomType)
                            {
                                // 最後尋找週數，週數通常比較固定，直接用 TryGetProperty
                                // 如果你的 JSON 裡面的週數也有大小寫問題，可以使用 targetWeek.ToUpper()
                                if (room.Value.TryGetProperty(targetWeek, out var priceElement))
                                {
                                    tuitionPrice = priceElement.GetDecimal();
                                    break; // 找到了，跳出房型迴圈
                                }
                                else if (room.Value.TryGetProperty(targetWeek.ToLower(), out priceElement))
                                {
                                    // 有些 JSON 寫 "4w"，有些寫 "4W"，雙重保險
                                    tuitionPrice = priceElement.GetDecimal();
                                    break;
                                }
                            }
                        }
                    }

                    // 如果已經找到了，就提早跳出外層迴圈，節省效能
                    if (tuitionPrice > 0) break;
                }
                if (tuitionPrice > 0) break;
            }

            // 3. 如果還是找不到，把前端傳來的參數印出來看看到底差在哪
            if (tuitionPrice == 0)
            {
                return BadRequest($"在 JSON 中找不到對應價格。您搜尋的條件為：課程[{request.Course}], 房型[{request.RoomType}], 週數[{targetWeek}]");
            }

            // 如果找不到對應的週數價格
            if (tuitionPrice == 0) return BadRequest("找不到對應的課程或週數價格");

            Tuple<decimal, decimal> retCalculationFee = tuitionFeeCalculate.CalculationFee("", request.Course, request.RoomType, request.Weeks, (int)tuitionPrice);

            // 3. 建立「課程費用項目」清單 (依照 R_163035.xlsx 的格式)
            CourseFeeItem courseFeeItem = new CourseFeeItem()  
            {
                Key = "1",
                Item = "註冊費",
                Content = "辦理註冊入學費用",
                Weeks = $"{request.Weeks}週",
                People = 1,
                UnitPrice = ZoneData.appSettings.CourseFee.RegistrationFee,
                Amount = 0,
                Remark = "註冊完成不退還"
            };
            courseFeeItems.Add(courseFeeItem);//建立註冊費

            if (request.NeedGuardianFee)
            {
                courseFeeItem = new CourseFeeItem()
                {
                    Key = "2",
                    Item = "未成年管理費",
                    Content = $"US{ZoneData.appSettings.CourseFee.LegalGuardianFee}/4週",
                    Weeks = $"{request.Weeks}週",
                    People = 1,
                    UnitPrice = ZoneData.appSettings.CourseFee.LegalGuardianFee,
                    Amount = 0,
                    Remark = ""
                };
                courseFeeItems.Add(courseFeeItem);//建未成年管理費
            }

            courseFeeItem = new CourseFeeItem()
            {
                Key = "3",
                Item = "學費",
                Content = request.Course,
                Weeks = $"{request.Weeks}週",
                People = 1,
                UnitPrice = retCalculationFee.Item1.ToString(),
                Amount = 0,
                Remark = ""
            };
            courseFeeItems.Add(courseFeeItem); //建立學費

            courseFeeItem = new CourseFeeItem()
            {
                Key = "4",
                Item = "住宿費",
                Content = request.RoomType,
                Weeks = $"{request.Weeks}週",
                People = 1,
                UnitPrice = retCalculationFee.Item2.ToString(),
                Amount = 0,
                Remark = ""
            };
            courseFeeItems.Add(courseFeeItem);//建立住宿費

            decimal SummerSurcharge = tuitionFeeCalculate.CalculateSurchargeWeek(request.StartDate, request.EndDate);
            if (SummerSurcharge > 0)
            {
                decimal.TryParse(ZoneData.appSettings.CourseFee.SummerSurcharge, out decimal oneWeekFee);
                courseFeeItem = new CourseFeeItem()
                {
                    Key = "5",
                    Item = "暑期加價",
                    Content = $"{SummerSurcharge}週",
                    Weeks = $"{request.Weeks}週",
                    People = 1,
                    UnitPrice = Math.Ceiling(SummerSurcharge * oneWeekFee).ToString(),
                    Amount = 0,
                    Remark = ""
                };
                courseFeeItems.Add(courseFeeItem);//建未成年管理費
            }

            //計算代辦折扣
            decimal.TryParse(ZoneData.appSettings.CourseFee.DiscountInclusive, out decimal DiscountInclusive);
            decimal.TryParse(ZoneData.appSettings.CourseFee.Discount, out decimal DiscountPercent);
            decimal.TryParse(ZoneData.appSettings.CourseFee.FixedDiscount, out decimal FixedDiscount);
            double DiscountFee = (((double)tuitionPrice + (double)(DiscountInclusive) + (double)(request.SchoolDiscount)) * ((double)DiscountPercent)) * -1;
            if (FixedDiscount != 0) DiscountFee = (double)FixedDiscount;

            string AgentDiscntContent = $"(學費+住宿費)*{DiscountPercent}%折扣";
            if (request.SchoolDiscount < 0)
            {
                courseFeeItem = new CourseFeeItem()
                {
                    Key = "6",
                    Item = "學校折扣",
                    Content = $"(選填)",
                    Weeks = $"{request.Weeks}週",
                    People = 1,
                    UnitPrice = request.SchoolDiscount.ToString(),
                    Amount = 0,
                    Remark = ""
                };
                courseFeeItems.Add(courseFeeItem);//建學校折扣
                AgentDiscntContent = AgentDiscntContent.Replace(")*", "-學校折扣)*");
            }
            
            courseFeeItem = new CourseFeeItem()
            {
                Key = "7",
                Item = "代辦折扣",
                Content = AgentDiscntContent,
                Weeks = $"{request.Weeks}週",
                People = 1,
                UnitPrice = DiscountFee.ToString(),
                Amount = 0,
                Remark = ""
            };
            courseFeeItems.Add(courseFeeItem);//建學校折扣

            // 4. 計算總計
            decimal totalUSD = 100 + tuitionPrice; // 註冊費 + 學費住宿費

            //[加入Local Fee]
            path = $"{schoolName.ToUpper()}_LocalFeeData.json"; // 之後可以改成 $"{schoolName}_tuition.json"
            path = Path.Combine(_env.WebRootPath, "files", path);

            if (!System.IO.File.Exists(path)) return NotFound($"找不到學校 {schoolName} 的學雜費資料");
            jsonString = System.IO.File.ReadAllText(path);

            // 這裡使用動態解析 (JsonDocument) 是最簡單的方式，因為可以自動展開動態的 Key
            using JsonDocument doc2 = JsonDocument.Parse(jsonString);
            var root2 = doc2.RootElement;

            // 2. 遍歷 tuition_data.json 的結構提取資料
            // 結構：大分類(DORMITORY) -> 課程(Special Study Permit SSP) -> "1W": 7800
            string reqPlaceofStay = request.Placeofstay?.Trim().ToLower() ?? "";
            localFeeList.Clear();
            foreach (var category in root2.EnumerateObject())
            {
                string jsonPlaceofStayName = category.Name.Trim().ToLower();
                if (jsonPlaceofStayName == reqPlaceofStay)
                {
                    foreach (var item in category.Value.EnumerateObject())
                    {
                        string DocItem = item.Name;
                        Console.WriteLine($"DocItem: {DocItem}");

                        // 最後尋找週數，週數通常比較固定，直接用 TryGetProperty
                        // 如果你的 JSON 裡面的週數也有大小寫問題，可以使用 targetWeek.ToUpper()
                        if (item.Value.TryGetProperty(targetWeek, out var priceElement) ||
                            item.Value.TryGetProperty(targetWeek.ToLower(), out priceElement))
                        {
                            LocalFeeInfo localFeeInfo = ZoneData.appSettings.LocalFee.FirstOrDefault(localFee => localFee.DocItem == DocItem || localFee.DocItem.Contains(DocItem)); 

                            string Key = localFeeInfo.Code;
                            string Item = localFeeInfo.Item;
                            string Content = localFeeInfo.Content;
                            string UnitPrice = priceElement.ToString();
                            string Weeks = localFeeInfo.Times.Contains("次") ? localFeeInfo.Times : $"{request.Weeks}週";
                            string Remark = localFeeInfo.Remark;
                            courseFeeItem = new CourseFeeItem()
                            {
                                Key = Key,
                                Item = Item,
                                Content = Content,
                                Weeks = Weeks,
                                People = 1,
                                UnitPrice = UnitPrice,
                                Amount = 0,
                                Remark = Remark
                            };
                            localFeeList.Add(courseFeeItem);
                        }
                    }
                }
            }

            // 只比對有 Item 屬性的項目
            var localFeeItems = localFeeList
                .OfType<CourseFeeItem>()  // 只取 CourseFeeItem
                .Select(cfi => cfi.Item)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            List<LocalFeeInfo> missingLocalFeeInfos = ZoneData.appSettings.LocalFee
                .Where(z => !localFeeItems.Contains(z.Item))
                .ToList();
            foreach (LocalFeeInfo localFeeInfo in missingLocalFeeInfos)
            {
                string Key = localFeeInfo.Code;
                string Item = localFeeInfo.Item;
                string Content = localFeeInfo.Content;

                int.TryParse(localFeeInfo.Price, out int price);
                string UnitPrice = (price * request.Weeks).ToString();  
                string Weeks = localFeeInfo.Times.Contains("次") ? localFeeInfo.Times : $"{request.Weeks}週";
                string Remark = localFeeInfo.Remark;
                courseFeeItem = new CourseFeeItem()
                {
                    Key = Key,
                    Item = Item,
                    Content = Content,
                    Weeks = Weeks,
                    People = 1,
                    UnitPrice = UnitPrice,
                    Amount = 0,
                    Remark = Remark
                };
                localFeeList.Add(courseFeeItem);
            }

            return Ok(new
            {
                CourseFees = courseFeeItems,
                localFees = localFeeList,
                TotalUSD = totalUSD,
                TotalNTD = totalUSD * request.UsaExchangeRate, //美金匯率要改成來自富邦銀行
                otherFees = new List<object>
                {
                    new { key = "air", item = "來回機票",Content="NT $6,000~20,000不等",Weeks="宿霧", people = 1, unitPrice = request.AirTicket, amount = request.AirTicket, remark = "可依航班調整" },
                    new { key = "visa", item = "簽證費",Content="紙本NT$1,200/電子NT$1,500",Weeks="紙本", people = 1, unitPrice = request.Visa, amount = request.Visa, remark = "線上/紙本申請" },
                    new { key = "ins", item = "旅平險/醫療險",Content="依各保險公司定價",Weeks="", people = 1, unitPrice = request.Insurance, amount = request.Insurance, remark = "建議投保3個月" }
                }
            });
        }
    


        //[HttpPost("calculate")]
        //public IActionResult Calculate([FromBody] object request)
        //{
        //    operationExcel.GetLocalFeeData();

        //    operationExcel.GetTuitionData();

        //    //operationExcel.ParseLocalFeeToJson();

        //    //operationExcel.ParseTUITIONExcelToJson();

        //    operationExcel.ParseSchoolJson("Philinter");
        //    //
        //    //operationExcel.ParseAcademyTuition("");

        //    // 模擬計算，4週 = 4800 tuition + 5600 dorm
        //    var weeks = 4;
        //    var tuition = 1200m * weeks;
        //    var dormitory = 1400m * weeks;
        //    var localTotal = 7800 + (250m * weeks); // SSP + 水費
        //    var exchangeRate = 32.05m;
        //    var totalNT = (tuition + dormitory + localTotal) * exchangeRate + 12500; // +機票簽證險

        //    return Ok(new
        //    {
        //        Tuition = tuition,
        //        Dormitory = dormitory,
        //        LocalTotal = localTotal,
        //        FixedTotal = 12500m,
        //        TotalNT = totalNT,
        //        LocalFees = new[]
        //        {
        //            new { Item = "SSP", Amount = 7800m },
        //            new { Item = "Water fee", Amount = 250m * weeks },
        //            new { Item = "Electricity", Amount = 700m * weeks }
        //        }
        //    });
        //}

        [HttpPost("export")]
        public IActionResult Export()
        {
            return Ok(new { message = "匯出成功！(開發中)" });
        }
    }
}
