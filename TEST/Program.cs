TimeSpan godzRozpPracy = new TimeSpan(17, 0, 0);
TimeSpan godzZakonczPracy = new TimeSpan(6, 0, 0);
double godzNadlPlatne50 = 2;
double godzNadlPlatne100 = 6;

double czasPrzepracowany = 0;

if (godzZakonczPracy < godzRozpPracy)
{
    czasPrzepracowany = (TimeSpan.FromHours(24) - godzRozpPracy).TotalHours + godzZakonczPracy.TotalHours;
}
else
{
    czasPrzepracowany = (godzZakonczPracy - godzRozpPracy).TotalHours;
}

double czasPodstawowy = czasPrzepracowany - (godzNadlPlatne50 + godzNadlPlatne100);

TimeSpan startPodstawowy = godzRozpPracy;
TimeSpan endPodstawowy = startPodstawowy + TimeSpan.FromHours(czasPodstawowy);

TimeSpan startNadl50 = endPodstawowy;
TimeSpan endNadl50 = startNadl50 + TimeSpan.FromHours(godzNadlPlatne50);

TimeSpan startNadl100 = endNadl50;
TimeSpan endNadl100 = startNadl100 + TimeSpan.FromHours(godzNadlPlatne100);

TimeSpan FormatTimeSpan(TimeSpan time)
{
    return ;
}
return (new TimeSpan((int)startPodstawowy.TotalHours % 24, startPodstawowy.Minutes, startPodstawowy.Seconds)),

Console.WriteLine("Podstawowy: " + FormatTimeSpan(startPodstawowy) + " - " + FormatTimeSpan(endPodstawowy));
Console.WriteLine("Nadgodziny 50%: " + FormatTimeSpan(startNadl50) + " - " + FormatTimeSpan(endNadl50));
Console.WriteLine("Nadgodziny 100%: " + FormatTimeSpan(startNadl100) + " - " + FormatTimeSpan(endNadl100));
