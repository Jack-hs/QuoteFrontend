using OfficeOpenXml.Style;

public class ExcelStyleData
{
    public string RangeRow { get; set; } = string.Empty;
    public string RangeCol { get; set; } = string.Empty;
    public string FontName { get; set; } = "微軟正黑體";
    public int FontSize { get; set; } = 12;
    public bool FontBold { get; set; } = false;
    public bool WrapText { get; set; } = true;
    public System.Drawing.Color FontColor { get; set; } = System.Drawing.Color.Black;
    public System.Drawing.Color BackgroundColor { get; set; }= System.Drawing.Color.White;
    public ExcelBorderStyle BorderStyle { get; set; } = ExcelBorderStyle.None;
    public System.Drawing.Color BorderColor { get; set; } = System.Drawing.Color.Black;
    public ExcelHorizontalAlignment HorizontalAlignment { get; set; } = ExcelHorizontalAlignment.Center;
    public ExcelVerticalAlignment VerticalAlignment { get; set; } = ExcelVerticalAlignment.Center;
    public string[] RangeNumberformat { get; set; } = {"", "", "NT$#,##0" };
}