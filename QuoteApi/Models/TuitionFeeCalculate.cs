using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Text;

namespace QuoteApi.Models
{
    public class TuitionFeeCalculate
    {
        public decimal CalculateSurchargeWeek(DateTime startDate, DateTime endDate)
        {
            if (startDate > endDate) return 0;

            TimeSpan totalDays = endDate - startDate;
            int totalWeeks = (int)Math.Ceiling(totalDays.TotalDays / 7.0);  // 向上取整

            // 檢查是否落在加價期間（例如 8/2-8/15）


            int[] startArray = Array.ConvertAll(ZoneData.appSettings.CourseFee.SummerSurchargeStartTime.Split('/'), int.Parse);
            int[] endArray = Array.ConvertAll(ZoneData.appSettings.CourseFee.SummerSurchargeEndTime.Split('/'), int.Parse);
            DateTime surchargeStart = new DateTime(startDate.Year, startArray[0], startArray[1]);
            DateTime surchargeEnd = new DateTime(startDate.Year, endArray[0], endArray[1]);

            int surchargeWeeks = 0;

            // 計算重疊的完整週數
            DateTime current = startDate;
            while (current < endDate)
            {
                DateTime weekEnd = current.AddDays(6);
                if (weekEnd >= surchargeStart && current <= surchargeEnd)
                {
                    surchargeWeeks++;
                }
                current = weekEnd.AddDays(1);
            }

            return surchargeWeeks;
        }

        public Tuple<decimal, decimal> CalculationFee(string cB_Tuition, string cB_course, string cB_room, int weeks, int price)
        {
            Tuple<decimal, decimal> ret = new Tuple<decimal, decimal>(0, 0);//學費,房價
            //int.TryParse(Weeks, out int weeks);//週數
            //int.TryParse(Price, out int price);//總價


            //如果已經包含註冊費 要扣掉再算
            int.TryParse(ZoneData.appSettings.CourseFee.RegistrationFee, out int RegistrationFee);
            if (ZoneData.appSettings.CourseFee.RegistrationFeeInclusive.ToLower() != "no")
            {

                price -= RegistrationFee;
            }


            decimal TuitionFee = (decimal)ZoneData.appSettings.Courses
                .SelectMany(category => category.Value) // 將所有課程陣列攤平成單一集合
                .Where(course => course.Name == cB_course) // 篩選課程名稱
                .Select(course => course.PricePerWeek) // 取出每週價格
                .FirstOrDefault(); // 取得第一筆符合的結果，找不到則回傳預設值 (0)


            //decimal roomTypeFee = livePrice[cB_room];
            decimal roomTypeFee = (decimal)ZoneData.appSettings.Room.Where(room => room.Name == cB_room)
                    .Select(room => room.PricePerWeek).FirstOrDefault();

            decimal[] less4WeekValues = ZoneData.appSettings.Less4Week.Select(x => decimal.Parse(x.Value)).ToArray(); // 或使用 x.Value 取值並轉型
            decimal[] More4WeekValues = ZoneData.appSettings.More4Week.Select(x => decimal.Parse(x.Value)).ToArray(); // 或使用 x.Value 取值並轉型
            try
            {
                //double TotalPrice = CourseRoomsPairs[cB_course.Text][cB_room.Text][weeks - 1];
                int remainder = (weeks % 4);// 0.25, 0.5, 0.75
                int Quotient = (weeks / 4);// * (TuitionFee + roomTypeFee)

                decimal multiple = Quotient == 0 ? less4WeekValues[remainder] : More4WeekValues[remainder];

                decimal basePrice = (TuitionFee + roomTypeFee);

                ret = new Tuple<decimal, decimal>(Math.Ceiling(TuitionFee * (Quotient + multiple)),
                    Math.Ceiling(roomTypeFee * (Quotient + multiple)));// Fee * 0.25 or * 1.25

                if (Math.Abs((double)(Math.Ceiling(ret.Item1 + ret.Item2) - price)) > 1) Console.WriteLine("學費註冊費計算不對等");
            }
            catch (Exception)
            {

                throw;
            }

            return ret;
        }
    }
}
