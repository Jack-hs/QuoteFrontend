using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Caching.Memory;
using Microsoft.Office.Interop.Excel;
using NPOI.SS.Formula.Functions;
using OfficeOpenXml;
using OfficeOpenXml.Drawing; // 確保此 using 語句存在
using OfficeOpenXml.Style;   // 確保此 using 語句存在
using QuoteApi.Models; // 假設您的 QuoteExportDto 和相關模型定義在此命名空間
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace ExcelApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExportQuoteController : ControllerBase
    {
        private List<ExcelStyleData> _excelStyles;
        int[] RecordNtdRowIndex;
        private int rowIndex;
        private readonly IMemoryCache _cache;

        public ExportQuoteController(IMemoryCache cache)
        {
            _cache = cache;
        }
        private int RowHeight = 25;

        [HttpPost("from-file")]
        public IActionResult ExportQuote([FromForm] IFormFile quoteJson)
        {
            if (quoteJson == null || quoteJson.Length == 0)
                return BadRequest("無 JSON 檔案");

            try
            {
                // 1. 讀取檔案內容 → JSON 字串 → 反序列化
                using var reader = new StreamReader(quoteJson.OpenReadStream());
                string jsonContent = reader.ReadToEnd();

                var jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<QuoteExportDto>(jsonContent);
                if (jsonData == null) return BadRequest("JSON 格式錯誤");

                // 設置 EPPlus 的授權上下文 (非商業用途)
                // 對於商業用途，請替換為 ExcelPackage.LicenseContext = LicenseContext.Commercial;
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                //OfficeOpenXml.LicenseManager.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                //ExcelPackage.License = new License(LicenseContext.NonCommercial);
                // 修正: 移除所有程式碼中的授權設定。
                // 授權已在 Program.cs 中設定，確保在 ExcelPackage 實例化之前完成。
                //ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                var templatePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "templates", "Sample.xlsx");
                ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization"); //This will also set the Company property to the organization name provided in the argument.
                using (var package = new ExcelPackage(templatePath))
                // 雙重保險：在 ExcelPackage 建構子
                //using (var package = new ExcelPackage(OfficeOpenXml.LicenseContext.NonCommercial))
                {
                    _excelStyles = new List<ExcelStyleData>();
                    if (_cache.TryGetValue("LatestRates", out ExchangeRates rates))
                    {
                        jsonData.basicInfo.usdRate = (double)rates.USD.SellRate;
                        jsonData.basicInfo.phpRate = (double)rates.PHP.SellRate;
                    }
                    else
                    {
                        return BadRequest("沒有匯率資料");
                    }
                    
                    //var worksheet = package.Workbook.Worksheets.Add("Sheet1");
                    var worksheet = package.Workbook.Worksheets[0];

                    //插入學校Logo
                    AddPicture(worksheet, jsonData.basicInfo.school);

                    // 填充數據 (根據您的 ClosedXML 邏輯進行調整)
                    // 注意：EPPlus 的儲存格索引是從 1 開始，與 ClosedXML 類似
                    worksheet.Cells["I10"].Value = DateTime.Now.ToString("yyyy/MM/dd");
                    worksheet.Cells["I11"].Value = $"{jsonData.basicInfo.studentName}";

                    RecordNtdRowIndex = new int[] { 25, 38, 45 }; // 台幣換算、披索換算、總金額的行索引
                    rowIndex = 15; // 從第15行開始填寫資料
                    worksheet.Cells[$"A{rowIndex}"].Value = $"{jsonData.basicInfo.school}";
                    // ... 其他數據填充邏輯，將 ClosedXML 的 ws.Cell 替換為 worksheet.Cells
                    // 例如：worksheet.Cells[$"C{rowIndex}"].Value = $"{ZoneData.appSettings.Setting.SchoolLocation}";
                    // 由於 ZoneData.appSettings 未提供，這裡僅作示意
                    worksheet.Cells[$"C{rowIndex}"].Value = $"{ZoneData.appSettings.Setting.SchoolLocation}";
                    worksheet.Cells[$"D{rowIndex}"].Value = $"{jsonData.basicInfo.course}";
                    // 假設 roomType 邏輯已在其他地方處理或直接提供
                    string roomType = ZoneData.appSettings.Room.
                                    Where(room => room.Name == jsonData.basicInfo.roomType).
                                    Select(room => room.Description).FirstOrDefault().ToString();
                    worksheet.Cells[$"F{rowIndex}"].Value = $"{roomType}";
                    worksheet.Cells[$"H{rowIndex}"].Value = $"{jsonData.basicInfo.weeks}";
                    worksheet.Cells[$"I{rowIndex}"].Value = $"{jsonData.basicInfo.startDate.Replace("-", "/")}- {jsonData.basicInfo.endDate.Replace("-", "/")}";


                    QuoteStructure quoteStructure = new QuoteStructure();
                    quoteStructure.headerCourseFee = new string[] { "課程費用項目", "費用內容", "週數", "人數", "單價", "金額(美金)", "備註" };
                    quoteStructure.jsonData = jsonData;
                    quoteStructure.FeeItem = quoteStructure.jsonData.CourseFees;
                    quoteStructure.StatementStr1 = "以上費用包含課程、住宿、餐食";
                    quoteStructure.ForeignTotalStr = "美金合計";
                    quoteStructure.textboxContentStr = "繳給語宙，幣別可台幣OR美金";
                    quoteStructure.StatementStr2 = $" 美金：台幣＝1：{quoteStructure.jsonData.basicInfo.usdRate}(報價當日匯率)";
                    quoteStructure.NtdTotalStr = "台幣換算";
                    quoteStructure.RangeNumberformat = "\"US\"#,##0";
                    quoteStructure.Rate = quoteStructure.jsonData.basicInfo.usdRate;
                    FillExcelCellData(worksheet, quoteStructure, 0, RowHeight);


                    quoteStructure = new QuoteStructure();
                    quoteStructure.headerCourseFee = new string[] { "當地雜費項目", "費用內容", "週數", "人數", "單價", "金額(披索)", "備註" };
                    quoteStructure.jsonData = jsonData;
                    quoteStructure.FeeItem = quoteStructure.jsonData.LocalFees;
                    quoteStructure.StatementStr1 = "以上費用實際以學校收取為主";
                    quoteStructure.ForeignTotalStr = "披索合計";
                    quoteStructure.textboxContentStr = "到當地用披索繳給學校";
                    quoteStructure.StatementStr2 = $"披索：台幣＝1：{jsonData.basicInfo.phpRate}(報價當日匯率)";
                    quoteStructure.NtdTotalStr = "台幣換算";
                    quoteStructure.RangeNumberformat = "\\₱#,##0";
                    quoteStructure.Rate = quoteStructure.jsonData.basicInfo.phpRate;
                    FillExcelCellData(worksheet, quoteStructure, 1, RowHeight);


                    quoteStructure = new QuoteStructure();
                    quoteStructure.headerCourseFee = new string[] { "機票簽證費用", "費用內容", "", "人數", "單價", "金額(台幣)", "備註" };
                    quoteStructure.jsonData = jsonData;
                    quoteStructure.FeeItem = quoteStructure.jsonData.OtherFees;
                    quoteStructure.StatementStr1 = "以上費用為預估費用，可請我們代訂代辦也可自行處理";
                    quoteStructure.ForeignTotalStr = "台幣合計";
                    quoteStructure.textboxContentStr = "預估遊學總費用(不含生活費及旅遊費用)";
                    quoteStructure.StatementStr2 = $"";
                    quoteStructure.NtdTotalStr = "台幣換算";
                    quoteStructure.RangeNumberformat = "\"NT$\"#,##0";
                    quoteStructure.Rate = 1;
                    FillExcelCellData(worksheet, quoteStructure, 2, RowHeight);


                    #region MyRegion
                    ////// 接下來處理課程費用
                    //rowIndex += 3; // 跳過3行 標題列18
                    //string[] headerCourseFee = new string[] { "課程費用項目", "費用內容", "週數", "人數", "單價", "金額(美金)", "備註" };
                    //for (int i = 0; i < headerCourseFee.Length; i++)
                    //{
                    //    // 這裡需要根據您的實際排版調整列索引
                    //    // 假設對應的列是 1, 2, 4, 5, 6, 7, 9
                    //    int[] headerCols = { 1, 2, 4, 5, 6, 7, 9 };
                    //    if (i < headerCols.Length)
                    //    {
                    //        worksheet.Cells[rowIndex, headerCols[i]].Value = headerCourseFee[i];
                    //        worksheet.Cells[rowIndex, headerCols[i]].Style.Font.Bold = true; // 設置標題粗體
                    //    }
                    //}
                    //ExcelStyleData excelStyleData = new ExcelStyleData(); //[課程費用項目]-標頭列
                    //excelStyleData.RangeRow = $"A{rowIndex}";
                    //excelStyleData.RangeCol = $"I{rowIndex}";
                    //excelStyleData.FontColor = Color.White;
                    //excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                    //excelStyleData.BorderStyle = ExcelBorderStyle.None;
                    //_excelStyles.Add(excelStyleData);

                    //var courseFees = jsonData.CourseFees;
                    //rowIndex += 1; // 跳到資料列19
                    //excelStyleData = new ExcelStyleData(); //[課程費用項目]-費用項目內容
                    //excelStyleData.RangeRow = $"A{rowIndex}";
                    //foreach (var fee in courseFees)
                    //{
                    //    worksheet.Cells[rowIndex, 1].Value = fee.item;              // 註冊費
                    //    worksheet.Cells[rowIndex, 2].Value = fee.content;           // 說明
                    //    worksheet.Cells[rowIndex, 4].Value = fee.weeks;             // 週數
                    //    worksheet.Cells[rowIndex, 5].Value = fee.people;            // 1
                    //    worksheet.Cells[rowIndex, 6].Value = fee.unitPrice;      // USD單價
                    //    // EPPlus 中設置公式
                    //    // 注意：公式中的儲存格引用應與實際的 rowIndex 和 column 匹配
                    //    worksheet.Cells[rowIndex, 7].Formula = $"F{rowIndex}*E{rowIndex}"; // 假設 F 列是單價，E 列是人數
                    //    worksheet.Cells[rowIndex, 9].Value = fee.remark; // 備註
                    //    rowIndex++;
                    //}
                    //excelStyleData.RangeCol = $"I{rowIndex}";
                    //excelStyleData.FontColor = Color.Black;
                    //excelStyleData.BackgroundColor = Color.White;
                    //excelStyleData.BorderStyle = ExcelBorderStyle.Medium;
                    //excelStyleData.RangeNumberformat = new string[] { $"{excelStyleData.RangeRow.Replace("A", "F")}", $"{excelStyleData.RangeCol.Replace("I", "G")}", "\"US\"#,##0" };
                    //_excelStyles.Add(excelStyleData);

                    ////跳到資料列24
                    //// 以上費用包含課程、住宿、餐食
                    //worksheet.Cells[rowIndex, 1].Value = "以上費用包含課程、住宿、餐食";
                    //worksheet.Cells[rowIndex, 4].Value = "美金合計";
                    //worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - courseFees.Count}:G{rowIndex - 1})";
                    ////AddStyledTextBox(ws, "繳給語宙，幣別可台幣OR美金", RowIndex, 9, 9);
                    //string textboxContent = "繳給語宙，幣別可台幣OR美金";
                    //AddFloatingShape(worksheet, "1", textboxContent, rowIndex, 8, 250, 35);
                    //excelStyleData = new ExcelStyleData(); //[課程費用項目]-以上費用包含課程、住宿、餐食
                    //excelStyleData.RangeRow = $"A{rowIndex}";
                    //excelStyleData.RangeCol = $"C{rowIndex}";
                    //excelStyleData.FontColor = Color.White;
                    //excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                    //excelStyleData.BorderStyle = ExcelBorderStyle.None;
                    //excelStyleData.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_excelStyles.Add(excelStyleData);

                    //excelStyleData = new ExcelStyleData(); //[課程費用項目]-美金合計
                    //excelStyleData.RangeRow = $"D{rowIndex}";
                    //excelStyleData.RangeCol = $"I{rowIndex}";
                    //excelStyleData.FontColor = Color.White;
                    //excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                    //excelStyleData.BorderStyle = ExcelBorderStyle.None;
                    //_excelStyles.Add(excelStyleData);


                    //rowIndex += 1; // 跳到資料列25
                    //worksheet.Cells[rowIndex, 1].Value = $" 美金：台幣＝1：{jsonData.basicInfo.usdRate}(報價當日匯率)";
                    //worksheet.Cells[rowIndex, 4].Value = "台幣換算";
                    //worksheet.Cells[rowIndex, 10].Value = $"{jsonData.basicInfo.usdRate}";
                    //worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - 1}*J{rowIndex})";
                    //RecordNtdRowIndex[0] = rowIndex; // 記錄課程費用台幣換算的行索引

                    //excelStyleData = new ExcelStyleData(); //[課程費用項目]-美金：台幣＝1：{jsonData.basicInfo.usdRate}(報價當日匯率)
                    //excelStyleData.RangeRow = $"A{rowIndex}";
                    //excelStyleData.RangeCol = $"C{rowIndex}";
                    //excelStyleData.FontColor = Color.White;
                    //excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                    //excelStyleData.BorderStyle = ExcelBorderStyle.None;
                    //excelStyleData.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    //_excelStyles.Add(excelStyleData);

                    //excelStyleData = new ExcelStyleData(); //[課程費用項目]-台幣換算
                    //excelStyleData.RangeRow = $"D{rowIndex}";
                    //excelStyleData.RangeCol = $"I{rowIndex}";
                    //excelStyleData.FontColor = Color.White;
                    //excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                    //excelStyleData.BorderStyle = ExcelBorderStyle.None;
                    //_excelStyles.Add(excelStyleData);

                    ////// 接下來處理當地雜費
                    //headerCourseFee = new string[] { "當地雜費項目", "費用內容", "週數", "人數", "單價", "金額(披索)", "備註" };
                    //rowIndex += 2; // 跳到資料列27
                    //for (int i = 0; i < headerCourseFee.Length; i++)
                    //{
                    //    // 這裡需要根據您的實際排版調整列索引
                    //    // 假設對應的列是 1, 2, 4, 5, 6, 7, 9
                    //    int[] headerCols = { 1, 2, 4, 5, 6, 7, 9 };
                    //    if (i < headerCols.Length)
                    //    {
                    //        worksheet.Cells[rowIndex, headerCols[i]].Value = headerCourseFee[i];
                    //        worksheet.Cells[rowIndex, headerCols[i]].Style.Font.Bold = true; // 設置標題粗體
                    //    }
                    //}
                    //rowIndex += 1; // 跳到資料列28
                    //var LocalFee = jsonData.LocalFees;
                    //foreach (var fee in LocalFee)
                    //{
                    //    worksheet.Cells[rowIndex, 1].Value = fee.item;              // 註冊費
                    //    worksheet.Cells[rowIndex, 2].Value = fee.content;           // 說明
                    //    worksheet.Cells[rowIndex, 4].Value = fee.weeks;             // 週數
                    //    worksheet.Cells[rowIndex, 5].Value = fee.people;            // 1
                    //    worksheet.Cells[rowIndex, 6].Value = fee.unitPrice;      // USD單價
                    //    // EPPlus 中設置公式
                    //    // 注意：公式中的儲存格引用應與實際的 rowIndex 和 column 匹配
                    //    worksheet.Cells[rowIndex, 7].Formula = $"F{rowIndex}*E{rowIndex}"; // 假設 F 列是單價，E 列是人數
                    //    worksheet.Cells[rowIndex, 9].Value = fee.remark; // 備註
                    //    rowIndex++;
                    //}
                    ////跳到資料列37
                    //// 以上費用包含課程、住宿、餐食
                    //worksheet.Cells[rowIndex, 1].Value = "以上費用實際以學校收取為主";
                    //worksheet.Cells[rowIndex, 4].Value = "披索合計";
                    //worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - LocalFee.Count}:G{rowIndex - 1})";
                    ////AddStyledTextBox(ws, "繳給語宙，幣別可台幣OR美金", RowIndex, 9, 9);
                    //textboxContent = "到當地用披索繳給學校";
                    //AddFloatingShape(worksheet, "2", textboxContent, rowIndex, 8, 200, 35);
                    //rowIndex += 1; // 跳到資料列38
                    //worksheet.Cells[rowIndex, 1].Value = $"披索：台幣＝1：{jsonData.basicInfo.phpRate}(報價當日匯率)";
                    //worksheet.Cells[rowIndex, 4].Value = "披索換算";
                    //worksheet.Cells[rowIndex, 10].Value = $"{jsonData.basicInfo.phpRate}";
                    //worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - 1}*J{rowIndex})";
                    //RecordNtdRowIndex[1] = rowIndex; // 記錄當地雜費台幣換算的行索引


                    ////// 接下來處理其他費用
                    //headerCourseFee = new string[] { "機票簽證費用", "費用內容", "", "人數", "單價", "金額(台幣)", "備註" };
                    //rowIndex += 2; // 跳到資料列40
                    //for (int i = 0; i < headerCourseFee.Length; i++)
                    //{
                    //    // 這裡需要根據您的實際排版調整列索引
                    //    // 假設對應的列是 1, 2, 4, 5, 6, 7, 9
                    //    int[] headerCols = { 1, 2, 4, 5, 6, 7, 9 };
                    //    if (i < headerCols.Length)
                    //    {
                    //        worksheet.Cells[rowIndex, headerCols[i]].Value = headerCourseFee[i];
                    //        worksheet.Cells[rowIndex, headerCols[i]].Style.Font.Bold = true; // 設置標題粗體
                    //    }
                    //}
                    //rowIndex += 1; // 跳到資料列41
                    //var OtherFee = jsonData.OtherFees;
                    //foreach (var fee in OtherFee)
                    //{
                    //    worksheet.Cells[rowIndex, 1].Value = fee.item;              // 註冊費
                    //    worksheet.Cells[rowIndex, 2].Value = fee.content;           // 說明
                    //    worksheet.Cells[rowIndex, 4].Value = fee.weeks;             // 週數
                    //    worksheet.Cells[rowIndex, 5].Value = fee.people;            // 1
                    //    worksheet.Cells[rowIndex, 6].Value = fee.unitPrice;      // USD單價
                    //    // EPPlus 中設置公式
                    //    // 注意：公式中的儲存格引用應與實際的 rowIndex 和 column 匹配
                    //    worksheet.Cells[rowIndex, 7].Formula = $"F{rowIndex}*E{rowIndex}"; // 假設 F 列是單價，E 列是人數
                    //    worksheet.Cells[rowIndex, 9].Value = fee.remark; // 備註
                    //    rowIndex++;
                    //}
                    ////跳到資料列44
                    //// 以上費用為預估費用，可請我們代訂代辦也可自行處理
                    //worksheet.Cells[rowIndex, 1].Value = "以上費用為預估費用，可請我們代訂代辦也可自行處理";
                    //worksheet.Cells[rowIndex, 4].Value = "台幣合計";
                    //worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - OtherFee.Count}:G{rowIndex - 1})";
                    //RecordNtdRowIndex[2] = rowIndex; // 記錄其他費用台幣的行索引
                    #endregion

                    var range = worksheet.Cells[$"A{rowIndex - 1}:C{rowIndex}"];
                    // 強制重新 merge
                    range.Merge = false;
                    range.Merge = true;

                    rowIndex += 1; //跳到資料列45
                    StatementCellRowSetting(worksheet, RowHeight, "G", "H");
                    ExcelStyleData excelStyleData = new ExcelStyleData(); //總金額
                    excelStyleData.RangeRow = $"A{rowIndex}";
                    excelStyleData.RangeCol = $"I{rowIndex}";
                    excelStyleData.FontColor = System.Drawing.Color.White;
                    excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(185, 121, 91);
                    excelStyleData.RangeNumberformat = new string[] { $"G{rowIndex}", $"G{rowIndex}", "\"NT$\"#,##0" };
                    _excelStyles.Add(excelStyleData);
                    worksheet.Cells[rowIndex, 6].Value = "總金額";
                    worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{RecordNtdRowIndex[0]} + G{RecordNtdRowIndex[1]} + G{RecordNtdRowIndex[2]})";
                    

                    rowIndex += 2; //跳到資料列47
                    excelStyleData = new ExcelStyleData(); //報價須知 :
                    excelStyleData.RangeRow = $"A{rowIndex}";
                    excelStyleData.RangeCol = $"I{rowIndex}";
                    excelStyleData.FontColor = System.Drawing.Color.FromArgb(44, 69, 27);
                    excelStyleData.FontBold = true;
                    excelStyleData.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    _excelStyles.Add(excelStyleData);
                    worksheet.Cells[rowIndex, 1].Value = "報價須知 :";

                    //報價須知內容
                    foreach (var Term in ZoneData.QuotationTermsDict)
                    {
                        rowIndex++;
                        StatementCellRowSetting(worksheet, 18, "A", "I");
                        excelStyleData = new ExcelStyleData();
                        excelStyleData.RangeRow = $"A{rowIndex}";
                        excelStyleData.RangeCol = $"I{rowIndex}";
                        excelStyleData.FontBold = true;
                        excelStyleData.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                        _excelStyles.Add(excelStyleData);
                        worksheet.Cells[rowIndex, 1].Value = $"{Term.Key}";
                        //string termText = string.Join("\n", Term.Value.Select(tuple => worksheet.Cells[tuple.Item1, tuple.Item2].Text));

                        if (Term.Value.Count() > 0)
                        {
                            int currentPos = 0;
                            var cell = worksheet.Cells[rowIndex, 1];
                            cell.Value = Term.Key;  // 完整文字
                            cell.RichText.Clear();

                            foreach (var tuple in Term.Value)
                            {
                                string partText = Term.Key.Substring(currentPos, tuple.Item1 - currentPos);
                                var part = cell.RichText.Add(partText);
                                part.Bold = true;
                                part.Color = System.Drawing.Color.Black;

                                string highlightText = Term.Key.Substring(tuple.Item1, tuple.Item2);
                                var highlightPart = cell.RichText.Add(highlightText);
                                highlightPart.Bold = true;
                                highlightPart.Color = System.Drawing.Color.Red;
                                currentPos = tuple.Item1 + tuple.Item2;
                            }
                        }
                    }

                    rowIndex++;
                    StatementCellRowSetting(worksheet, 18, "A", "I");
                    excelStyleData = new ExcelStyleData(); //最後一行綠色停止線
                    excelStyleData.RangeRow = $"A{rowIndex}";
                    excelStyleData.RangeCol = $"I{rowIndex}";
                    excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(44, 69, 27);
                    _excelStyles.Add(excelStyleData);

                    //最外面的灰色邊框
                    range = worksheet.Cells[$"A7:I{rowIndex}"];
                    range.Style.Border.BorderAround(ExcelBorderStyle.Medium, System.Drawing.Color.FromArgb(200, 200, 200));

                    // 在所有數據填充完成後，應用樣式
                    RenderExcelStyleData(worksheet);

                    //// 將 Excel 文件保存到記憶體流中
                    //var stream = new MemoryStream();
                    //package.SaveAs(stream);
                    //stream.Position = 0;

                    //// 返回 Excel 文件作為檔案流
                    //return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"報價單_{jsonData.basicInfo.school}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");

                    //1.指定儲存路徑
                    string fileName = $"報價單_{jsonData.basicInfo.school}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
                    string savePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "exports", fileName);

                    //// 2. 建立資料夾（如果不存在）
                    //Directory.CreateDirectory(Path.GetDirectoryName(savePath));

                    //// 3. 直接儲存檔案
                    //package.SaveAs(savePath);

                    //// 4. 回傳檔案路徑或成功訊息
                    //return Ok(new
                    //{
                    //    message = "Excel 已儲存本地",
                    //    filePath = savePath,
                    //    downloadUrl = $"/exports/{fileName}"
                    //});
                    using var stream = new MemoryStream();
                    package.SaveAs(stream);
                    stream.Position = 0;

                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"報價單_{jsonData.basicInfo.school}.xlsx");
                }
            }
            catch (Exception ex)
            {
                // 記錄詳細錯誤信息，便於調試
                Console.WriteLine($"處理錯誤: {ex.Message}\n{ex.StackTrace}");
                return StatusCode(500, $"處理錯誤: {ex.Message}\n{ex.StackTrace}");
            }
        }

        private class QuoteStructure
        {
            public string[] headerCourseFee { get; set; }
            public QuoteExportDto jsonData { get; set; }
            public List<FeeItem> FeeItem { get; set; }
            public string StatementStr1 { get; set; }
            public string ForeignTotalStr { get; set; }
            public string textboxContentStr { get; set; }
            public string StatementStr2 { get; set; }
            public string NtdTotalStr { get; set; }
            public string RangeNumberformat { get; set; }
            public double Rate { get; set; }
            
        }

        private void FillExcelCellData(ExcelWorksheet worksheet, QuoteStructure quoteStructure, int index, int RowHeight)
        {
            try
            {
                //// 接下來處理課程費用
                rowIndex += 3; // 跳過3行 標題列18
                string[] headerCourseFee = quoteStructure.headerCourseFee;
                DataCellRowSetting(worksheet, RowHeight);
                for (int i = 0; i < headerCourseFee.Length; i++)
                {
                    // 這裡需要根據您的實際排版調整列索引
                    // 假設對應的列是 1, 2, 4, 5, 6, 7, 9
                    int[] headerCols = { 1, 2, 4, 5, 6, 7, 9 };
                    if (i < headerCols.Length)
                    {
                        worksheet.Cells[rowIndex, headerCols[i]].Value = headerCourseFee[i];
                        worksheet.Cells[rowIndex, headerCols[i]].Style.Font.Bold = true; // 設置標題粗體
                    }
                }
                ExcelStyleData excelStyleData = new ExcelStyleData(); //[課程費用項目]-標頭列
                excelStyleData.RangeRow = $"A{rowIndex}";
                excelStyleData.RangeCol = $"I{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.White;
                excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                excelStyleData.BorderStyle = ExcelBorderStyle.None;
                _excelStyles.Add(excelStyleData);
                

                var courseFees = quoteStructure.FeeItem;
                foreach (var fee in courseFees)
                {

                    rowIndex++;

                    int TempRowHeight = RowHeight;
                    if (fee.remark.Count() > 20) TempRowHeight = RowHeight + ((fee.remark.Count() / 20) * 5);
                    
                    DataCellRowSetting(worksheet, TempRowHeight);
                    worksheet.Cells[rowIndex, 1].Value = fee.item;              // 註冊費
                    worksheet.Cells[rowIndex, 2].Value = fee.content;           // 說明
                    worksheet.Cells[rowIndex, 4].Value = fee.weeks;             // 週數
                    worksheet.Cells[rowIndex, 5].Value = fee.people;            // 1
                    worksheet.Cells[rowIndex, 6].Value = fee.unitPrice;      // USD單價
                                                                             // EPPlus 中設置公式
                                                                             // 注意：公式中的儲存格引用應與實際的 rowIndex 和 column 匹配
                    worksheet.Cells[rowIndex, 7].Formula = $"F{rowIndex}*E{rowIndex}"; // 假設 F 列是單價，E 列是人數
                    worksheet.Cells[rowIndex, 9].Value = fee.remark; // 備註
                }
                excelStyleData = new ExcelStyleData(); //[課程費用項目]-費用項目內容
                excelStyleData.RangeRow = $"A{rowIndex - courseFees.Count + 1}";
                excelStyleData.RangeCol = $"I{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.Black;
                excelStyleData.BackgroundColor = System.Drawing.Color.White;
                excelStyleData.BorderStyle = ExcelBorderStyle.Thin;
                excelStyleData.BorderColor = System.Drawing.Color.Black;
                excelStyleData.RangeNumberformat = new string[] { $"{excelStyleData.RangeRow.Replace("A", "F")}", $"{excelStyleData.RangeCol.Replace("I", "G")}", quoteStructure.RangeNumberformat};
                _excelStyles.Add(excelStyleData);

                excelStyleData = new ExcelStyleData(); //[課程費用項目]-備註改紅字
                excelStyleData.RangeRow = $"I{rowIndex - courseFees.Count + 1}";
                excelStyleData.RangeCol = $"I{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.Red;
                _excelStyles.Add(excelStyleData);

                // 以上費用包含課程、住宿、餐食
                rowIndex += 1; //跳到資料列24
                StatementCellRowSetting(worksheet, RowHeight, "A", "C");
                StatementCellRowSetting(worksheet, RowHeight, "G", "H");
                worksheet.Cells[rowIndex, 1].Value = quoteStructure.StatementStr1;
                worksheet.Cells[rowIndex, 6].Value = quoteStructure.ForeignTotalStr;
                worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - courseFees.Count}:G{rowIndex - 1})";
                //AddStyledTextBox(ws, "繳給語宙，幣別可台幣OR美金", RowIndex, 9, 9);
                string textboxContent = quoteStructure.textboxContentStr;
                AddFloatingShape(worksheet, DateTime.Now.ToString("ssff"), textboxContent, rowIndex, 8, 250, 35);
                excelStyleData = new ExcelStyleData(); //[課程費用項目]-以上費用包含課程、住宿、餐食
                excelStyleData.RangeRow = $"A{rowIndex}";
                excelStyleData.RangeCol = $"C{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.White;
                excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                excelStyleData.BorderStyle = ExcelBorderStyle.None;
                excelStyleData.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                _excelStyles.Add(excelStyleData);

                excelStyleData = new ExcelStyleData(); //[課程費用項目]-美金合計
                excelStyleData.RangeRow = $"D{rowIndex}";
                excelStyleData.RangeCol = $"I{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.White;
                excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                excelStyleData.BorderStyle = ExcelBorderStyle.None;
                excelStyleData.RangeNumberformat = new string[] { $"G{rowIndex}", $"H{rowIndex}", quoteStructure.RangeNumberformat };
                _excelStyles.Add(excelStyleData);


                rowIndex += 1; // 跳到資料列25
                StatementCellRowSetting(worksheet, RowHeight, "A", "C");
                StatementCellRowSetting(worksheet, RowHeight, "G", "H");
                worksheet.Cells[rowIndex, 1].Value = quoteStructure.StatementStr2;
                worksheet.Cells[rowIndex, 6].Value = quoteStructure.NtdTotalStr;
                worksheet.Cells[rowIndex, 10].Value = $"{quoteStructure.Rate}";
                worksheet.Cells[rowIndex, 7].Formula = $"=sum(G{rowIndex - 1}*J{rowIndex})";
                RecordNtdRowIndex[index] = rowIndex; // 記錄課程費用台幣換算的行索引

                if (quoteStructure.NtdTotalStr == "")
                {
                    worksheet.Cells[$"A{rowIndex - 1}:C{rowIndex}"].Merge = true;
                }

                excelStyleData = new ExcelStyleData(); //[課程費用項目]-美金：台幣＝1：{jsonData.basicInfo.usdRate}(報價當日匯率)
                excelStyleData.RangeRow = $"A{rowIndex}";
                excelStyleData.RangeCol = $"C{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.White;
                excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(163, 138, 118);
                excelStyleData.BorderStyle = ExcelBorderStyle.None;
                excelStyleData.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                _excelStyles.Add(excelStyleData);

                excelStyleData = new ExcelStyleData(); //[課程費用項目]-台幣換算
                excelStyleData.RangeRow = $"D{rowIndex}";
                excelStyleData.RangeCol = $"I{rowIndex}";
                excelStyleData.FontColor = System.Drawing.Color.White;
                excelStyleData.BackgroundColor = System.Drawing.Color.FromArgb(131, 60, 12);
                excelStyleData.BorderStyle = ExcelBorderStyle.None;
                excelStyleData.RangeNumberformat = new string[] { $"G{rowIndex}", $"H{rowIndex}", "\"NT$\"#,##0" };
                _excelStyles.Add(excelStyleData);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                throw;
            }
        }

        private void RenderExcelStyleData(ExcelWorksheet worksheet)
        {
            // 在這裡實現 ExcelStyleData 的樣式應用邏輯
            // 例如，遍歷 _excelStyles 列表，根據每個 ExcelStyleData 的屬性設置對應的儲存格樣式
            // 這裡需要根據您的 ExcelStyleData 定義來實現具體的樣式應用邏輯
            foreach (var styleData in _excelStyles)
            {
                //根據 styleData 的屬性設置儲存格樣式
                //例如：
                var range = worksheet.Cells[$"{styleData.RangeRow}:{styleData.RangeCol}"];
                range.Style.Font.Name = styleData.FontName;
                range.Style.Font.Size = styleData.FontSize;
                range.Style.Font.Bold = styleData.FontBold;
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Font.Color.SetColor(styleData.FontColor);
                range.Style.Fill.BackgroundColor.SetColor(styleData.BackgroundColor);
                if (styleData.BorderStyle != ExcelBorderStyle.None)
                {
                    //range.Style.Border.BorderAround(styleData.BorderStyle, styleData.BorderColor);
                    // 方法 2：分別設定上下左右邊框
                    range.Style.Border.Top.Style = styleData.BorderStyle;
                    range.Style.Border.Top.Color.SetColor(styleData.BorderColor);

                    range.Style.Border.Bottom.Style = styleData.BorderStyle;
                    range.Style.Border.Bottom.Color.SetColor(styleData.BorderColor);

                    range.Style.Border.Left.Style = styleData.BorderStyle;
                    range.Style.Border.Left.Color.SetColor(styleData.BorderColor);

                    range.Style.Border.Right.Style = styleData.BorderStyle;
                    range.Style.Border.Right.Color.SetColor(styleData.BorderColor);
                }
                range.Style.HorizontalAlignment = styleData.HorizontalAlignment;
                range.Style.VerticalAlignment = styleData.VerticalAlignment;
                if (styleData.WrapText) range.Style.WrapText = true;
                if (styleData.RangeNumberformat[0] != "")
                {
                    range = worksheet.Cells[$"{styleData.RangeNumberformat[0]}:{styleData.RangeNumberformat[1]}"];
                    range.Style.Numberformat.Format = styleData.RangeNumberformat[2];
                }
            }
        }

        private void AddFloatingShape(ExcelWorksheet worksheet, string Name, string Text, int Row, int Col, int Width, int Height, int UpOffset = -20, int eightOffset = -20)
        {
            if (Text == "") return;
            var shape = worksheet.Drawings.AddShape($"txtDesc{Name}_{Guid.NewGuid()}", eShapeStyle.Rect);
            //shape.SetPosition(1, 5, 6, 5);  //Position Row, RowOffsetPixels, Column, ColumnOffsetPixels
            //shape.SetSize(400, 200);        //Size in pixels
            shape.SetPosition(Row, UpOffset, Col, eightOffset);
            shape.SetSize(Width, Height);
            shape.EditAs = eEditAs.Absolute;
            shape.Text = $"{Text}";
            shape.Fill.Style = eFillStyle.SolidFill;
            shape.Fill.Color = System.Drawing.Color.FromArgb(218, 227, 243);
            shape.Fill.Transparancy = 0;
            shape.Border.Fill.Color = System.Drawing.Color.FromArgb(68, 114, 196);
            shape.Border.Width = 3f;
            shape.TextAnchoring = eTextAnchoringType.Center;
            shape.TextVertical = eTextVerticalType.Horizontal;
            shape.TextAnchoringControl = false;
            shape.TextAlignment = eTextAlignment.Center;
            shape.Font.Color = System.Drawing.Color.Black;
            shape.Font.Bold = true;

            shape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterRight);
            shape.Effect.SetPresetGlow(ePresetExcelGlowType.Accent3_8Pt);
        }

        private void DataCellRowSetting(ExcelWorksheet worksheet, int RowHeight)
        {
            worksheet.Cells[$"B{rowIndex}:C{rowIndex}"].Merge = true;
            worksheet.Cells[$"G{rowIndex}:H{rowIndex}"].Merge = true;

            worksheet.Row(rowIndex).Height = RowHeight;
            worksheet.Row(rowIndex).CustomHeight = true;  // 重要！
        }

        private void StatementCellRowSetting(ExcelWorksheet worksheet, int RowHeight, string StartCol, string EndCcol)
        {
            worksheet.Cells[$"{StartCol}{rowIndex}:{EndCcol}{rowIndex}"].Merge = true;
            

            worksheet.Row(rowIndex).Height = RowHeight;
            worksheet.Row(rowIndex).CustomHeight = true;  // 重要！
        }

        private void AddPicture(ExcelWorksheet worksheet,string schoolName)
        {
            var imagePath = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "AcademyLogo", $"{schoolName}.png");
            var image = new FileInfo(imagePath);

            // 加入圖片
            var picture = worksheet.Drawings.AddPicture($"Logo_{Guid.NewGuid()}", image);

            // 設定位置（row, rowOffset, col, colOffset）
            int.TryParse(ZoneData.appSettings.Logo.PicLocationLeft, out int PicLocationLeft);
            int.TryParse(ZoneData.appSettings.Logo.PicLocationTop, out int PicLocationTop);

            int.TryParse(ZoneData.appSettings.Logo.PicLocationRow, out int PicLocationRow);
            int.TryParse(ZoneData.appSettings.Logo.PicLocationCol, out int PicLocationCol);
            picture.SetPosition(PicLocationRow, PicLocationLeft, PicLocationCol, PicLocationTop);


            int.TryParse(ZoneData.appSettings.Logo.PicWidth, out int PicWidth);
            int.TryParse(ZoneData.appSettings.Logo.PicHeight, out int PicHeight);
            // 設定大小
            picture.SetSize(PicWidth, PicHeight); // width, height (px)
        }

        public void AddFloatingTextBox(ExcelWorksheet worksheet, string Name, string Text, int Row, int Col, int Width, int Height, int UpOffset = -20, int eightOffset = -20)
        {
            //// 插入文字框的邏輯
            //// 這裡的 left, top, width, height 參數是像素值
            //int textboxLeft = UpOffset; // 距離左邊緣的像素
            //int textboxTop = eightOffset;   // 距離頂部的像素
            //int textboxWidth = Width; // 寬度像素
            //int textboxHeight = Height; // 高度像素
            //string textboxContent = Text;

            ////// EPPlus 8.5.0 中，AddTextBox 方法返回 ExcelShape
            ////// ✅ 原子計數器生成，絕對唯一
            ////string uniqueName = DrawingNameGenerator.GetUniqueName($"TextBox_{Name}");
            ////Console.WriteLine($"AddFloatingTextBox uniqueName: {uniqueName}");
            ////// ✅ 雙重保險：預檢查
            ////if (worksheet.Drawings[uniqueName] != null)
            ////{
            ////    throw new InvalidOperationException($"意外！名稱 {uniqueName} 已存在");
            ////}
            ////OfficeOpenXml.Drawing.ExcelShape textBox = worksheet.Drawings.AddTextbox(uniqueName, string.Empty);


            //// 🔥 關鍵：傳空字串，讓 EPPlus 自動生成唯一名稱！
            //var textBox = worksheet.Drawings.AddTextbox("", Text);  // ✅ 空名稱 = 自動唯一

            //Console.WriteLine($"自動生成名稱：{textBox.Name}");  // 會顯示如 "TextBox 1"

            //textBox.SetPosition(Row, textboxLeft, Col, textboxTop);
            //textBox.SetSize(textboxWidth, textboxHeight);

            //// 強制重新設定文字（因為初始 Text 可能有問題）
            //textBox.RichText.Clear();


            //// SetPosition 的參數類型和重載。使用基於行/列的偏移量。
            //// SetPosition(int fromRow, int fromCol, int fromRowOffsetPixels, int fromColOffsetPixels)
            ////textBox.SetPosition(Row, textboxLeft, Col, textboxTop);
            ////textBox.SetSize(textboxWidth, textboxHeight);

            //// 修正: 根據 EPPlus 8.5.0 API，RichText.Add() 返回 ExcelParagraph
            //// 並且 Font 屬性在 ExcelParagraph 的 Style 屬性下。
            //OfficeOpenXml.Style.ExcelParagraph paragraph = textBox.RichText.Add(textboxContent);
            //paragraph.SetFromFont("微軟正黑體", 12, true, false, false, false); // 範例：字型, 大小, bold, italic, underline, strikeout
            //paragraph.Color = System.Drawing.Color.Black;
            //textBox.TextAlignment = eTextAlignment.Center;


            //// 如果有第二行文字，同樣處理
            //// OfficeOpenXml.Style.ExcelParagraph paragraph2 = textBox.RichText.Add("\n第二行文字");
            //// paragraph2.Style.Font.Bold = true;
            //// paragraph2.Style.Font.Color = System.Drawing.Color.Blue; // 範例

            //// 其他樣式設定
            //textBox.Fill.Color = Color.FromArgb(218, 227, 243);


            //textBox.Border.Fill.Color = Color.FromArgb(68, 114, 196);
            //textBox.Border.Width = 3f; // 邊框寬度
        }
    }
}



//using ClosedXML.Excel;
//using DocumentFormat.OpenXml;
//using DocumentFormat.OpenXml.Packaging;
//using DocumentFormat.OpenXml.Spreadsheet;
//using Microsoft.AspNetCore.Mvc;
//using Newtonsoft.Json.Linq;
//using OfficeOpenXml;
//using OfficeOpenXml.Drawing;
//using QuoteApi.Models;
//using static System.Runtime.InteropServices.JavaScript.JSType;

//[ApiController]
//[Route("api/[controller]")]
//public class ExportQuoteController : ControllerBase
//{

//    [HttpPost("from-file")]
//    public IActionResult ExportQuote([FromForm] IFormFile quoteJson)  // 改名避免混淆
//    {
//        if (quoteJson == null || quoteJson.Length == 0)
//            return BadRequest("無 JSON 檔案");

//        try
//        {
//            // ✅ 1. 讀取檔案內容 → JSON 字串 → 反序列化
//            using var reader = new StreamReader(quoteJson.OpenReadStream());
//            string jsonContent = reader.ReadToEnd();

//            var jsonData = Newtonsoft.Json.JsonConvert.DeserializeObject<QuoteExportDto>(jsonContent);
//            if (jsonData == null) return BadRequest("JSON 格式錯誤");

//            // ✅ 2. 以下你的 ClosedXML 邏輯（已完美）
//            var templatePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "templates", "Sample.xlsx");
//            using var workbook = new XLWorkbook(templatePath);
//            var ws = workbook.Worksheet(1);

//            ws.Cell("I10").Value = DateTime.Now.ToString("yyyy/MM/dd");
//            ws.Cell("I11").Value = $"Name";

//            int RowIndex = 15; // 從第15行開始填寫資料
//            ws.Cell($"A{RowIndex}").Value = $"{jsonData.basicInfo.school}";
//            ws.Cell($"C{RowIndex}").Value = $"{ZoneData.appSettings.Setting.SchoolLocation}";
//            ws.Cell($"D{RowIndex}").Value = $"{jsonData.basicInfo.course}";
//            string roomType = ZoneData.appSettings.Room.
//                Where(room => room.Name == jsonData.basicInfo.roomType).
//                Select(room => room.Description).FirstOrDefault().ToString();
//            ws.Cell($"F{RowIndex}").Value = $"{roomType}";
//            ws.Cell($"H{RowIndex}").Value = $"{jsonData.basicInfo.weeks}";
//            ws.Cell($"I{RowIndex}").Value = $"{jsonData.basicInfo.startDate.Replace("-", "/")}- {jsonData.basicInfo.endDate.Replace("-", "/")}";
//            // ... 你其他填資料邏輯完全保留

//            RowIndex += 3; // 跳過3行 標題列18
//            string[] HeaderCourseFee = new string[] { "課程費用項目", "費用內容", "週數", "人數", "單價", "金額(美金)", "備註" };
//            ws.Cell(RowIndex, 1).Value = HeaderCourseFee[0];
//            ws.Cell(RowIndex, 2).Value = HeaderCourseFee[1];
//            ws.Cell(RowIndex, 4).Value = HeaderCourseFee[2];
//            ws.Cell(RowIndex, 5).Value = HeaderCourseFee[3];
//            ws.Cell(RowIndex, 6).Value = HeaderCourseFee[4];
//            ws.Cell(RowIndex, 7).Value = HeaderCourseFee[5];
//            ws.Cell(RowIndex, 9).Value = HeaderCourseFee[6];

//            var courseFees = jsonData.CourseFees;
//            RowIndex += 1; // 跳到資料列19
//            for (int i = 0; i < courseFees.Count; i++)
//            {
//                RowIndex += 1;
//                ws.Cell(RowIndex, 1).Value = courseFees[i].item;              // 註冊費
//                ws.Cell(RowIndex, 2).Value = courseFees[i].content;           // 說明
//                ws.Cell(RowIndex, 4).Value = courseFees[i].weeks;             // 週數
//                ws.Cell(RowIndex, 5).Value = courseFees[i].people;            // 1
//                ws.Cell(RowIndex, 6).Value = courseFees[i].unitPrice;      // USD單價
//                ws.Cell(RowIndex, 7).FormulaA1 = $"=D{5}*E{6}";          // 小計
//                ws.Cell(RowIndex, 9).Value = courseFees[i].remark; // 備註
//            }

//            RowIndex += 1; // 跳到資料列24
//            // 以上費用包含課程、住宿、餐食
//            ws.Cell(RowIndex, 1).Value = "以上費用包含課程、住宿、餐食";
//            ws.Cell(RowIndex, 4).Value = "美金合計";
//            ws.Cell(RowIndex, 7).FormulaA1 = $"=sum(G{RowIndex - courseFees.Count}:G{RowIndex - 1})";
//            //AddStyledTextBox(ws, "繳給語宙，幣別可台幣OR美金", RowIndex, 9, 9);

//            RowIndex += 1; // 跳到資料列25
//            ws.Cell(RowIndex, 1).Value = $" 美金：台幣＝1：{jsonData.basicInfo.usdRate}(報價當日匯率)";
//            ws.Cell(RowIndex, 4).Value = "台幣換算";
//            ws.Cell(RowIndex, 10).Value = $"{jsonData.basicInfo.usdRate}";
//            ws.Cell(RowIndex, 7).FormulaA1 = $"=sum(G{RowIndex - 1}*J{RowIndex})";


//            // 1. 指定儲存路徑
//            string fileName = $"報價單_{jsonData.basicInfo.school}_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";
//            string savePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "exports", fileName);

//            // 2. 建立資料夾（如果不存在）
//            Directory.CreateDirectory(Path.GetDirectoryName(savePath));

//            // 3. 直接儲存檔案
//            workbook.SaveAs(savePath);

//            // 4. 回傳檔案路徑或成功訊息
//            return Ok(new
//            {
//                message = "Excel 已儲存本地",
//                filePath = savePath,
//                downloadUrl = $"/exports/{fileName}"
//            });
//            //using var stream = new MemoryStream();
//            //workbook.SaveAs(stream);
//            //stream.Position = 0;

//            //return File(stream.ToArray(),
//            //    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
//            //    $"報價單_{jsonData.basicInfo.school}.xlsx");
//        }
//        catch (Exception ex)
//        {
//            return BadRequest($"處理錯誤: {ex.Message}");
//        }
//    }


//}
