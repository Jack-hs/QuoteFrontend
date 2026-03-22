using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Http.HttpResults;
using System.Data;
//using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;
using static QuoteApi.Models.LocalFeeData;
using static QuoteApi.Models.TuitionData;
using Excel = Microsoft.Office.Interop.Excel;

namespace QuoteApi.Models
{
    public class OperationExcel
    {
        public void ParseLocalFeeToJson()
        {
            // 請替換為您的實際檔案路徑
            //string filePath = @"C:\Users\yhl25\source\repos\QuoteSystemOri\QuoteSystem\QuoteApi\PHILINTER-LOCAL-FEE-updated-on-November-26-2025.xlsx";
            //string outputPath = "local_fee_data.json";
            string fileName = "PHILINTER LOCAL FEE (updated on November 26, 2025).xlsx";
            string filePath = Path.Combine(ZoneData.StartupPath, "files", $"{fileName}");
            string outputName = "PHILINTER_LocalFeeData.json";
            string outputPath = Path.Combine(ZoneData.StartupPath, "files", $"{outputName}");

            // 註冊編碼提供者 (在 .NET Core 讀取 Excel 必需)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            // 第一層: 住宿類型 (DORMITORY / AZON)
            // 第二層: 收費項目 (PARTICULAR)
            // 第三層: 週數 -> 金額 (1W: 7800)
            var localFeeData = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet();

                    foreach (DataTable table in result.Tables)
                    {
                        // 簡化 Sheet 名稱 (如 "AZON - Condominium" 轉為 "AZON")
                        string sheetName = table.TableName.Replace(" - Condominium", "").Trim();
                        var sheetData = new Dictionary<string, Dictionary<string, double>>();

                        int headerRowIdx = -1;
                        int particularColIdx = -1;

                        // 1. 尋找包含 "PARTICULAR" 的標題列與欄位索引
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            for (int c = 0; c < table.Columns.Count; c++)
                            {
                                if (table.Rows[r][c]?.ToString()?.Trim() == "PARTICULAR")
                                {
                                    headerRowIdx = r;
                                    particularColIdx = c;
                                    break;
                                }
                            }
                            if (headerRowIdx != -1) break;
                        }

                        if (headerRowIdx == -1) continue; // 若找不到 PARTICULAR 標題列則跳過該頁籤

                        // 2. 從標題列的下一行開始讀取資料
                        for (int r = headerRowIdx + 1; r < table.Rows.Count; r++)
                        {
                            var row = table.Rows[r];
                            string particularName = CleanExcelString(row[particularColIdx])?.ToString()?.Trim();

                            // 過濾無效資料或多餘標題列
                            if (string.IsNullOrEmpty(particularName) ||
                                particularName == "Tuition" ||
                                particularName == "Accommodatioin" ||
                                particularName.Contains("DORMITORY") ||
                                particularName.Contains("AZON CONDO"))
                            {
                                continue;
                            }

                            // 碰到下方簽證說明的備註文字時，直接跳過
                            if (particularName.StartsWith("for ") || particularName.Contains("weeks"))
                            {
                                continue;
                            }
                            if (particularName.StartsWith("TOTAL AMOUNT"))
                            {
                                break;
                            }

                            var pricing = new Dictionary<string, double>();

                            // 3. 讀取右側所有的週數費用
                            for (int c = particularColIdx + 1; c < table.Columns.Count; c++)
                            {
                                string weekHeader = table.Rows[headerRowIdx][c]?.ToString()?.Trim();

                                // 讀取金額並清除千分位逗號
                                string priceStr = CleanExcelString(row[c])?.ToString()?.Replace(",", "").Trim();

                                // 確認標題是 "1W", "2W"... 且價格欄位有數字
                                if (!string.IsNullOrEmpty(weekHeader) && weekHeader.EndsWith("W"))
                                {
                                    if (double.TryParse(priceStr, out double price))
                                    {
                                        pricing[weekHeader] = price;
                                    }
                                }
                            }

                            // 若該項目有讀取到任何週數費用，則存入 Dictionary
                            if (pricing.Count > 0)
                            {
                                // 若項目名稱有換行符號（如 SSP (換行) ACR E-CARD），將其替換為空格以保持 JSON 乾淨
                                particularName = particularName.Replace("\n", " ").Replace("\r", "");
                                sheetData[particularName] = pricing;
                            }
                        }

                        if (sheetData.Count > 0)
                        {
                            localFeeData[sheetName] = sheetData;
                        }
                    }
                }
            }

            // 4. 輸出並序列化為 JSON
            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(localFeeData, options);
            File.WriteAllText(outputPath, jsonString);

            Console.WriteLine("Local Fee Excel 已成功轉換為 JSON 檔！");
        }

        public void GetLocalFeeData()
        {
            // 1. 讀取 JSON 檔案
            var fileName = Path.Combine(ZoneData.StartupPath, "files", "local_fee_data.json");
            if (!System.IO.File.Exists(fileName))
            {
                //return NotFound("找不到當地費用檔案");
                Console.WriteLine($"{fileName}.json 檔案不存在");
            }

            var jsonString = System.IO.File.ReadAllText(fileName);

            // 2. 反序列化
            var localFeeData = JsonSerializer.Deserialize<LocalFeeRoot>(jsonString);

            if (localFeeData == null)
            {
                //return BadRequest("當地費用資料格式錯誤");
            }

            // ==========================================
            // 3. 把資料存入對應的變數裡面 (全部都是字典)
            // ==========================================

            // 取得 DORMITORY 大分類的字典
            if (localFeeData.TryGetValue("DORMITORY", out var dormitoryFees))
            {
                // 變數 dormitoryFees 現在是一個字典，包含了所有 DORMITORY 下的雜費項目

                // 取得 "Special Study Permit (SSP)" 費用的字典
                if (dormitoryFees.TryGetValue("Special Study Permit (SSP)", out var sspFees))
                {
                    // 取得 SSP 費用 4週 (4W) 的價格
                    var sspPrice4W = sspFees.ContainsKey("4W") ? sspFees["4W"] : 0; // 7800
                }

                // 取得 "Water fee" (水費) 的字典
                if (dormitoryFees.TryGetValue("Water fee", out var waterFees))
                {
                    // 取得水費 12週 (12W) 的價格
                    var waterPrice12W = waterFees.ContainsKey("12W") ? waterFees["12W"] : 0; // 3000
                }

                // 取得 "Visa Extension fee" (簽證延簽費) 的字典
                if (dormitoryFees.TryGetValue("Visa Extension fee", out var visaFees))
                {
                    // 注意：這項費用不是每週都有 (只有 9W 以上才有)
                    // 所以一定要用 ContainsKey 或 TryGetValue 來檢查，否則會報錯
                    if (visaFees.TryGetValue("12W", out var visaPrice))
                    {
                        // 取得到 12W 價格
                    }
                    else
                    {
                        // 1W ~ 8W 沒有這個費用
                    }
                }
            }

            // 回傳完整資料給前端
            //return Ok(localFeeData);
        }

        public void ParseTUITIONExcelToJson()
        {
            string fileName = "PHILINTER TUITION (updated on November 25, 2025).xlsx";
            string filePath = Path.Combine(ZoneData.StartupPath, "files", $"{fileName}");
            string outputName = "PHILINTER_TuitionData.json";
            string outputPath = Path.Combine(ZoneData.StartupPath, "files", $"{outputName}");

            // 註冊編碼提供者 (在 .NET Core 或 .NET 5+ 以上版本讀取 Excel 必需)
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

            var tuitionData = new Dictionary<string, Dictionary<string, Dictionary<string, Dictionary<string, double>>>>();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // 使用 ExcelDataReader 讀取
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    // 將整個 Excel 轉換為 DataSet (包含多個 DataTable，每個 Table 對應一個 Sheet)
                    var result = reader.AsDataSet();

                    foreach (DataTable table in result.Tables)
                    {
                        string mainCourse = table.TableName.Split(" 2026")[0].Trim();
                        var sheetData = new Dictionary<string, Dictionary<string, Dictionary<string, double>>>();

                        int headerRowIdx = -1;

                        // 尋找包含 "Course" 的標題列位置
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            for (int c = 0; c < table.Columns.Count; c++)
                            {
                                if (table.Rows[r][c]?.ToString()?.Trim() == "Course")
                                {
                                    headerRowIdx = r;
                                    break;
                                }
                            }
                            if (headerRowIdx != -1) break;
                        }

                        if (headerRowIdx == -1) continue; // 找不到標題列就跳過這個分頁

                        string currentSubCourse = "";

                        // 從標題列的下一行開始讀取資料
                        for (int r = headerRowIdx + 1; r < table.Rows.Count; r++)
                        {
                            var row = table.Rows[r];

                            // 對應欄位：索引 1 是小課程(B欄)，索引 2 是房型(C欄)
                            string colB = row.ItemArray.Length > 1 ? CleanExcelString(row[1])?.ToString()?.Trim() : "";
                            string roomType = row.ItemArray.Length > 2 ? CleanExcelString(row[2])?.ToString()?.Trim() : "";

                            if (roomType != "" && colB == "")
                            {
                                colB = row.ItemArray.Length > 1 ? CleanExcelString(row[0])?.ToString()?.Trim() : "";
                            }

                            if (string.IsNullOrEmpty(roomType) || roomType == "-") continue;

                            // 更新當前小課程
                            if (!string.IsNullOrEmpty(colB) && colB != mainCourse)
                            {
                                currentSubCourse = colB;
                            }

                            if (string.IsNullOrEmpty(currentSubCourse)) continue;

                            if (!sheetData.ContainsKey(currentSubCourse))
                            {
                                sheetData[currentSubCourse] = new Dictionary<string, Dictionary<string, double>>();
                            }

                            var pricing = new Dictionary<string, double>();

                            // 讀取各週費用 (從索引 3，也就是 D 欄開始找 1W, 2W...)
                            for (int c = 3; c < table.Columns.Count; c++)
                            {
                                string weekHeader = table.Rows[headerRowIdx][c]?.ToString()?.Trim();
                                string priceStr = CleanExcelString(row[c])?.ToString()?.Trim();

                                if (!string.IsNullOrEmpty(weekHeader) && weekHeader.EndsWith("W"))
                                {
                                    if (double.TryParse(priceStr, out double price))
                                    {
                                        pricing[weekHeader] = price;
                                    }
                                }
                            }

                            if (pricing.Count > 0)
                            {
                                sheetData[currentSubCourse][roomType] = pricing;
                            }
                        }

                        if (sheetData.Count > 0)
                        {
                            tuitionData[mainCourse] = sheetData;
                        }
                    }
                }
            }

            // 輸出 JSON
            var options = new JsonSerializerOptions { WriteIndented = true };
            string jsonString = JsonSerializer.Serialize(tuitionData, options);
            File.WriteAllText(outputPath, jsonString);
        }
        
        public void GetTuitionData()
        {
            // 1. 讀取 JSON 檔案
            var fileName = Path.Combine(ZoneData.StartupPath, "files", "tuition_data.json");
            if (!System.IO.File.Exists(fileName))
            {
                //return NotFound("找不到學費檔案");
                Console.WriteLine($"{fileName}.json 檔案不存在");
            }

            var jsonString = System.IO.File.ReadAllText(fileName);

            // 2. 反序列化 (直接轉成我們定義好的多層字典)
            var tuitionData = JsonSerializer.Deserialize<TuitionRoot>(jsonString);

            if (tuitionData == null)
            {
                //return BadRequest("學費資料格式錯誤");
            }

            // ==========================================
            // 3. 把資料存入對應的變數裡面 (全部都是字典)
            // ==========================================

            // 取得 IELTS 大分類底下的所有課程字典
            if (tuitionData.TryGetValue("IELTS", out var ieltsCourses))
            {
                // 變數 ieltsCourses 現在是一個字典，包含了 IELTS 相關的所有課程

                // 取得 "IELTS INTENSIVE" 這個特定課程的各房型字典
                if (ieltsCourses.TryGetValue("IELTS INTENSIVE", out var ieltsIntensiveRooms))
                {
                    // 變數 ieltsIntensiveRooms 是字典，包含了 Single, Double 等房型

                    // 取得 "Single" (單人房) 的每週價格字典
                    if (ieltsIntensiveRooms.TryGetValue("Single", out var singleRoomPrices))
                    {
                        // 變數 singleRoomPrices 是字典，包含了 "1W", "2W"... 等價格

                        // 💡 實際應用：直接抓出 4W 的價格
                        var priceFor4Weeks = singleRoomPrices["4W"]; // 得到 2600
                    }
                }
            }

            // 取得 ESL 分類的字典
            if (tuitionData.TryGetValue("ESL & SPEAKING", out var eslCourses))
            {
                // eslCourses 變數存放 ESL & SPEAKING 底下的資料
            }

            // 回傳完整資料給前端
            //return Ok(tuitionData);
        }

        //將原本ini 當換學校時的基本資料 讀取到全域變數appSettings之中方便調用
        public void ParseSchoolJson(string fileName)
        {
            JsonDocument jsonDoc = null;
            try
            {
                //string fileName = "Philinter";
                var jsonPath = Path.Combine(ZoneData.StartupPath, "files", $"{fileName}.json");
                if (!System.IO.File.Exists(jsonPath))
                {
                    Console.WriteLine($"{fileName}.json 檔案不存在");
                }

                var jsonContent = System.IO.File.ReadAllText(jsonPath);
                jsonDoc = JsonDocument.Parse(jsonContent);

                // 1. 先讀取 JSON 並反序列化成原本的 AppSettings 類別
                ZoneData.appSettings = JsonSerializer.Deserialize<AppSettings>(jsonContent);

                var BasePrice = ZoneData.appSettings.Room.Where(room => room.Name == "Single")
                    .Select(room => room.PricePerWeek).FirstOrDefault();

                // 假設 appSettings.courses 的型別為 Dictionary<string, List<Course>>
                var price = ZoneData.appSettings.Courses
                    .SelectMany(category => category.Value) // 將所有課程陣列攤平成單一集合
                    .Where(course => course.Name == "INTENSIVE ESL") // 篩選課程名稱
                    .Select(course => course.PricePerWeek) // 取出每週價格
                    .FirstOrDefault(); // 取得第一筆符合的結果，找不到則回傳預設值 (0)

                if (ZoneData.appSettings != null)
                {
                    // ==========================================
                    // 第一部分：單一物件，直接存入對應的變數
                    // ==========================================
                    var setting = ZoneData.appSettings.Setting;
                    var courseFee = ZoneData.appSettings.CourseFee;
                    var logo = ZoneData.appSettings.Logo;


                    // ==========================================
                    // 第二部分：有一個以上的陣列，全部轉成字典 (Dictionary)
                    // ==========================================

                    // 1. 雜費 (LocalFee) 轉字典：使用 Code 當作 Key (例如 "1", "2", "3")
                    Dictionary<string, LocalFeeInfo> localFeeDict = ZoneData.appSettings.LocalFee
                        .ToDictionary(fee => fee.DocItem, fee => fee);

                    // 2. 房間 (Room) 轉字典：使用 Code 當作 Key (例如 "Room1", "Room2")
                    Dictionary<string, RoomInfo> roomDict = ZoneData.appSettings.Room
                        .ToDictionary(room => room.Name, room => room);

                    // 3. 課程 (Courses) 轉字典：
                    // 原本的結構是 Dictionary<分類名稱, List<課程>>
                    // 下面這段程式碼會把它變成 Dictionary<分類名稱, Dictionary<課程Code, 課程資訊>>
                    Dictionary<string, Dictionary<string, CourseInfo>> coursesDict = ZoneData.appSettings.Courses
                        .ToDictionary(
                            category => category.Key, // 外層的 Key 是分類名稱 (例如 "IELTS")
                            category => category.Value.ToDictionary(course => course.Name, course => course) // 內層的 Key 是課程 Code (例如 "IELTS1")
                        );


                    // ==========================================
                    // 第三部分：原本 JSON 就是 Key-Value 結構的資料
                    // ==========================================
                    Dictionary<string, string> more4WeekDict = ZoneData.appSettings.More4Week;
                    Dictionary<string, string> less4WeekDict = ZoneData.appSettings.Less4Week;


                    // 測試取值範例：
                    // 取出 Code 為 "Room2" 的房間名稱
                    var myRoomName = roomDict["AZON Double"].PricePerWeek; // 會得到 "Double"

                    // 取出 IELTS 分類下，Code 為 "IELTS2" 的課程名稱
                    //var myCourseName = coursesDict["JUNIOR & FAMILY"]["JUNIOR IELTS"].PricePerWeek; // 會得到 "IELTS GUARANTEE 12 WEEKS"
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"讀取設定失敗: {ex.Message}");
                throw;
            }
        }

        // 1. 建立一個輔助方法，專門用來清洗 Excel 讀出來的字串
        public string CleanExcelString(object cellValue)
        {
            

            if (cellValue == null || cellValue == DBNull.Value) return "";

            string rawString = cellValue.ToString().Trim();
            //return rawString;

            if (string.IsNullOrEmpty(rawString)) return "";

            // 步驟 A：將所有換行符號 (\n 或 \r\n 或 \r) 取代為一個空格
            string cleaned = rawString.Replace("\r\n", " ")
                                        .Replace("\n", " ")
                                        .Replace("\r", " ");

            // 2. 🌟 取代容易造成 JSON 或 URL 錯誤的特殊字元 🌟

            // (A) 針對特定符號進行「意義上」的替換
            cleaned = cleaned.Replace("&", "and");  // 例如: "ESL & SPEAKING" 變成 "ESL and SPEAKING"
            cleaned = cleaned.Replace("~", "-");    // 例如: "7~11yrs" 變成 "7-11yrs"

            // (B) 移除可能造成 JSON 破壞的引號
            cleaned = cleaned.Replace("\"", "");    // 移除雙引號
            cleaned = cleaned.Replace("'", "");     // 移除單引號

            // 3. 🌟 移除所有中文字 🌟
            // [\u4e00-\u9fa5] 是 Unicode 中用來配對「基本漢字」的範圍
            cleaned = Regex.Replace(cleaned, @"[\u4e00-\u9fa5]", "");
            // (C) 如果你想「只保留英文、數字、中文、括號、減號與空白」，把其他怪怪的符號全刪掉：
            // \u4e00-\u9fa5 代表中文
            // a-zA-Z0-9 代表英文數字
            // \(\)\[\]\- 代表括號與減號
            cleaned = Regex.Replace(cleaned, @"[\(\)\[\]\-]", "");
            // \s 代表空白
            // 如果你不需要這麼極端，可以把這行註解掉。這行非常強大，會把所有不在名單上的符號清空。
            // cleaned = Regex.Replace(cleaned, @"[^a-zA-Z0-9\u4e00-\u9fa5\(\)\[\]\-\s\.]", "");

            // 3. 把連續兩個以上的空白縮減為一個
            cleaned = Regex.Replace(cleaned, @"\s+", " ");

            return cleaned.Trim();
        }

        List<string> sheets = new List<string>();
        public System.Data.DataTable ParseAcademyTuition(string filePath, int headerRow = 1, int startRow = 2)
        {
            //JsonDocument jsonDoc = null;
            //try
            //{
            //    string fileName = "Philinter";
            //    var jsonPath = Path.Combine(ZoneData.StartupPath, "files", $"{fileName}.json");
            //    if (!System.IO.File.Exists(jsonPath))
            //    {
            //        Console.WriteLine($"{fileName}.json 檔案不存在");
            //    }

            //    var jsonContent = System.IO.File.ReadAllText(jsonPath);
            //    jsonDoc = JsonDocument.Parse(jsonContent);

            //    Console.WriteLine("Json read finish");
            //}
            //catch (Exception ex)
            //{
            //    Console.WriteLine($"讀取設定失敗: {ex.Message}");
            //    throw;
            //}

            Excel.Application xlApp = null;
            Excel.Workbook xlWorkbook = null;
            Excel.Worksheet xlWorksheet = null;
            Excel.Range xlRange = null;
            System.Data.DataTable dt = new System.Data.DataTable();

            Dictionary<string, List<string>> LocalItemPrice = new Dictionary<string, List<string>>();

            try
            {
                xlApp = new Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(filePath);

                //讀取頁籤
                sheets = new List<string>();
                foreach (Excel.Worksheet sheet in xlWorkbook.Sheets)
                {
                    sheets.Add(sheet.Name);
                }
                // MessageBox 選擇頁籤
                string selectedSheet = "";
                xlWorksheet = (Excel.Worksheet?)xlWorkbook.Sheets[1];  // 第一張工作表
                xlRange = xlWorksheet.Range["B9:Z20"]; //ini
                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                //1.從 headerRow 讀欄位名稱
                for (int c = 1; c <= colCount; c++)
                {
                    var headerCell = xlRange.Cells[headerRow, c];
                    string colName = headerCell?.ToString() ?? $"Column{c}";
                    dt.Columns.Add(colName, typeof(string));
                }
                Console.WriteLine("header finish");
                for (int r = startRow; r <= rowCount; r++)
                {
                    DataRow dr = dt.NewRow();

                    bool hasData = false;
                    for (int c = 1; c <= colCount; c++)
                    {
                        object cellValue = xlRange.Cells[r, c];
                        dr[c - 1] = cellValue ?? DBNull.Value;
                        string CellValue = dr[c - 1].ToString().Trim();

                        if (cellValue != null)
                            hasData = true;
                    }

                    // 只新增有資料的列
                    if (hasData)
                    {
                        dt.Rows.Add(dr);
                        string a = "";
                        string Value = "";
                        string Key = "";
                        foreach (var dr2 in dr.ItemArray)
                        {
                            a += dr2?.ToString() ?? "";
                            a += ",";

                            Value = dr2?.ToString().Trim() ?? "";
                            if (HasEnglishLetter(dr2?.ToString()))
                            {
                                Key = Value;
                                if (!LocalItemPrice.ContainsKey(Key))
                                {
                                    LocalItemPrice.Add(Key, new List<string>());
                                }
                            }
                            else
                            {
                                if (LocalItemPrice.ContainsKey(Key)) LocalItemPrice[Key].Add(Value);
                            }
                        }
                        Console.WriteLine($"r: {r}, cellValue: " + a);
                    }
                }
                Console.WriteLine("dt finish");
                ////讀ini 跟 dict 對照
                //TimeSpan duration = dTP_EndTime.Value.Date - dTP_StartTime.Value.Date;
                //int weeksFull = duration.Days / 7;
                //if ((duration.Days % 7) != 0) weeksFull += 1;
                //schoolSetting.DurWeeks = weeksFull.ToString();

                //string fileName = $"{cbB_schoolName.Text}.ini";
                //string iniPath = Path.Combine(System.Windows.Forms.Application.StartupPath, fileName);
                //if (!File.Exists(iniPath))
                //{
                //    var result = MessageBox.Show($"沒有相關文件檔", "警告", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //}
                //else
                //{
                //    //插入註冊費用 From ini
                //    string schoolNameIniPath = System.Windows.Forms.Application.StartupPath + $"//{cbB_schoolName.Text}.ini";
                //    SchoolName = new IniFile(schoolNameIniPath);
                //    string[] SchoolNames = SchoolName.GetKeys("LocalFeeItem");
                //    for (int i = 0; i <= SchoolNames.Count() - 1; i++)
                //    {
                //        Console.WriteLine($"i: {i}");
                //        string[] ItemDetail = SchoolName.IniReadUTF8("LocalFeeItem", SchoolNames[i]).Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                //        string FeeTimes = "";
                //        if (ItemDetail[3].Contains("次"))
                //        {
                //            FeeTimes = ItemDetail[3];
                //        }
                //        else if (ItemDetail[3].Contains("週"))
                //        {
                //            FeeTimes = $"{weeksFull}週";
                //        }

                //        string str = ""; int stri = 0;
                //        foreach (string s in ItemDetail)
                //        {
                //            str += $"ItemDetail[{stri}]: {s} --- "; stri += 1;
                //        }

                //        Console.WriteLine($"LocalFeeItem {str}");
                //        //if (LocalItemPrice.ContainsKey(ItemDetail[0]))

                //        Tuple<bool, string> ret = FindMatchingKey(LocalItemPrice.Keys.ToArray(), ItemDetail[0]);
                //        if (ret.Item1)
                //        {
                //            Console.WriteLine($"FindMatchingKey [True]: {ret.Item1} {ret.Item2}");
                //            //string[] ItemDetail = Item.Split(new char[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
                //            int weeksIndex = weeksFull > 0 ? weeksFull - 1 : 0;
                //            //if (LocalItemPrice[ItemDetail[0]].Count() < weeksIndex) return null;
                //            if (LocalItemPrice[ret.Item2].Count() < weeksIndex) return null;


                //            dataGridView2.Rows.Add($"{ItemDetail[1]}", $"{ItemDetail[2]}", $"{FeeTimes}", 1, $"{LocalItemPrice[ret.Item2][weeksIndex]}", "", $"{ItemDetail[5]}");
                //        }
                //        else
                //        {
                //            Console.WriteLine($"FindMatchingKey [No]: {ret.Item1} {ItemDetail[0]}");
                //            //option local fee
                //            int.TryParse(ItemDetail[4], out int oneWeekFee);
                //            dataGridView2.Rows.Add($"{ItemDetail[1]}", $"{ItemDetail[2]}", $"{FeeTimes}", 1, $"{oneWeekFee * weeksFull}", "", $"{ItemDetail[5]}");
                //        }
                //        //if (i == 0) continue;

                //        //cbB_schoolName.Items.Add(School);
                //    }

                //    //string registrationfee = schoolName.IniReadUTF8("CourseFee", "registrationfee");
                //    //dataGridView1.Rows.Add($"{"註冊費"}", $"{"辦理註冊入學費用"}", $"{Weeks}週", 1, registrationfee, "", "");
                //    //讀取INI LOGO 設定位置大小

                //    string PicLocationRow = SchoolName.IniReadUTF8("Logo", "PicLocationRow");
                //    if (!int.TryParse(PicLocationRow, out schoolSetting.PicLocationRow)) schoolSetting.PicLocationRow = 7;

                //    string PicLocationCol = SchoolName.IniReadUTF8("Logo", "PicLocationCol");
                //    if (!int.TryParse(PicLocationCol, out schoolSetting.PicLocationCol)) schoolSetting.PicLocationCol = 2;

                //    string PicLocationLeft = SchoolName.IniReadUTF8("Logo", "PicLocationLeft");
                //    if (!int.TryParse(PicLocationLeft, out schoolSetting.PicLocationLeft)) schoolSetting.PicLocationLeft = 0;

                //    string PicLocationTop = SchoolName.IniReadUTF8("Logo", "PicLocationTop");
                //    if (!int.TryParse(PicLocationTop, out schoolSetting.PicLocationTop)) schoolSetting.PicLocationTop = 0;

                //    string size = SchoolName.IniReadUTF8("Logo", "width");
                //    if (!int.TryParse(size, out schoolSetting.LogoWidth)) schoolSetting.LogoWidth = 110;

                //    size = SchoolName.IniReadUTF8("Logo", "height");
                //    if (!int.TryParse(size, out schoolSetting.LogoHeight)) schoolSetting.LogoHeight = 110;

                //    //SchoolName.IniReadUTF8
                //    //加入暑假加價 每週金額 時間
                //    string Surcharge = SchoolName.IniReadUTF8("CourseFee", "SummerSurcharge");
                //    if (!decimal.TryParse(Surcharge, out schoolSetting.SummerSurcharge)) schoolSetting.SummerSurcharge = 0;

                //    Surcharge = SchoolName.IniReadUTF8("CourseFee", "SummerSurchargeStartTime");
                //    if (Surcharge.Contains('/'))
                //    {
                //        string[] Value = Surcharge.Split('/');
                //        int.TryParse(Value[0], out int Month);
                //        int.TryParse(Value[1], out int Day);
                //        schoolSetting.SummerSurchargeStartTime = new DateTime(DateTime.Now.Year, Month, Day);
                //    }

                //    Surcharge = SchoolName.IniReadUTF8("CourseFee", "SummerSurchargeEndTime");
                //    if (Surcharge.Contains('/'))
                //    {
                //        string[] Value = Surcharge.Split('/');
                //        int.TryParse(Value[0], out int Month);
                //        int.TryParse(Value[1], out int Day);
                //        schoolSetting.SummerSurchargeEndTime = new DateTime(DateTime.Now.Year, Month, Day);
                //    }

                //    //判斷註冊有無包含在Tuition
                //    string registrationfeeInclusive = SchoolName.IniReadUTF8("CourseFee", "registrationfeeInclusive");
                //    schoolSetting.registrationfeeInclusive = registrationfeeInclusive.ToLower() == "no" ? false : true;

                //    //註冊費
                //    string registrationfee = SchoolName.IniReadUTF8("CourseFee", "registrationfee");
                //    if (!int.TryParse(registrationfee, out schoolSetting.registrationfee)) schoolSetting.registrationfee = 0;

                //    //代辦折扣包含的註冊費
                //    string DiscountInclusive = SchoolName.IniReadUTF8("CourseFee", "DiscountInclusive");
                //    if (!decimal.TryParse(DiscountInclusive, out schoolSetting.DiscountInclusive)) schoolSetting.DiscountInclusive = 0;

                //    //代辦折扣計算的乘數
                //    string Discount = SchoolName.IniReadUTF8("CourseFee", "Discount");
                //    if (!decimal.TryParse(Discount, out schoolSetting.Discount)) schoolSetting.Discount = 0;

                //    //代辦折扣 固定的
                //    string FixedDiscount = SchoolName.IniReadUTF8("CourseFee", "FixedDiscount");
                //    if (!decimal.TryParse(FixedDiscount, out schoolSetting.FixedDiscount)) schoolSetting.FixedDiscount = 0;

                //    //代辦折扣 固定的
                //    string LegalGuardianFee = SchoolName.IniReadUTF8("CourseFee", "LegalGuardianFee");
                //    if (!int.TryParse(LegalGuardianFee, out schoolSetting.LegalGuardianFee)) schoolSetting.LegalGuardianFee = 25;
                //    checkBox_LegalGuardianFee.Enabled = schoolSetting.LegalGuardianFee > 0 ? true : false;
                //}
                return dt;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ex: {ex}");
                throw;
            }
            finally
            {
                // 重要：釋放 COM 物件
                if (xlRange != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRange);
                if (xlWorksheet != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorksheet);
                if (xlWorkbook != null)
                {
                    xlWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlWorkbook);
                }
                if (xlApp != null)
                {
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);
                }
            }
        }

        public static bool HasEnglishLetter(string input)
        {
            return input.Any(c => Char.IsLetter(c) && (c >= 'A' && c <= 'Z' || c >= 'a' && c <= 'z'));
        }

        public static Tuple<bool, string> FindMatchingKey(string[] dictionaryKeys, string input)
        {
            foreach (string kvp in dictionaryKeys)
            {
                string value = kvp.Trim();
                if (kvp.Contains('&'))
                {
                    value = kvp.Split('&')[1];
                }

                if (CompareEnglishOnlyOrContains(input, value))
                {
                    return new Tuple<bool, string>(true, value);  // 返回對應英文
                }
            }
            return new Tuple<bool, string>(false, "");  // 沒找到
        }

        public static bool CompareEnglishOnlyOrContains(string str1, string str2)
        {
            bool ret = false;

            if (string.IsNullOrEmpty(str1) || string.IsNullOrEmpty(str2))
                return ret;

            // 只取 ASCII 英文字母 a-z A-Z → 小寫
            var english1 = (str1 ?? "").Where(c => c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z')
                                       .Select(c => char.ToLower(c));

            var english2 = (str2 ?? "").Where(c => c >= 'a' && c <= 'z' || c >= 'A' && c <= 'Z')
                                       .Select(c => char.ToLower(c));

            string newStr1 = "";
            foreach (var c in english1) newStr1 += c;

            string newStr2 = "";
            foreach (var c in english2) newStr2 += c;

            // 序列比較
            if (english1.SequenceEqual(english2)) ret = true;
            else
            {
                //if (newStr1.Contains(newStr2) || newStr2.Contains(newStr1)) ret = true;
                if (newStr2.Contains(newStr1)) ret = true;
            }

            return ret;
        }
    }
}
