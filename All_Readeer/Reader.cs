using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace All_Readeer
{
    internal static class Reader
    {
        public static TimeSpan Try_Get_Date(string cellValue)
        {
            if (TimeSpan.TryParse(cellValue, out TimeSpan time))
            {
                return time;
            }

            if (DateTime.TryParse(cellValue, out DateTime dateTime))
            {
                return dateTime.TimeOfDay;
            }

            if (TimeSpan.TryParse(cellValue.Replace(".", ":"), out TimeSpan timeWithColon))
            {
                return timeWithColon;
            }

            if (TimeSpan.TryParse(cellValue.Replace(";", ":"), out TimeSpan timeafterC))
            {
                return timeWithColon;
            }

            var parts = cellValue.Split(':');
            if (parts.Length == 3 &&
                int.TryParse(parts[0], out int hours) &&
                int.TryParse(parts[1], out int minutes) &&
                int.TryParse(parts[2], out int seconds))
            {
                if (seconds >= 60)
                {
                    seconds -= 60;
                    minutes++;
                }

                if (minutes >= 60)
                {
                    minutes -= 60;
                    hours++;
                }

                hours %= 24;
                return new TimeSpan(hours, minutes, seconds);
            }
            if (double.TryParse(cellValue, out double timeD))
            {
                return TimeSpan.FromDays(timeD);
            }


            throw new FormatException("Nieprawidłowy format godziny");
        }
        public static int Get_Month_Number_From_String(string input)
        {
            string[] months = new string[]
            {
            "styczeń", "luty", "marzec", "kwiecień", "maj", "czerwiec",
            "lipiec", "sierpień", "wrzesień", "październik", "listopad", "grudzień"
            };
            input = input.ToLower();

            for (int i = 0; i < months.Length; i++)
            {
                if (input.Contains(months[i]))
                {
                    return i + 1;
                }
            }

            return 0;
        }

    }
}
