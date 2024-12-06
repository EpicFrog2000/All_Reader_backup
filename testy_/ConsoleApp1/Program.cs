using ClosedXML.Excel;

//TODO errorlogger

namespace All_Readeer
{
    internal static class Grafik_Pracy_Reader_2024_v2
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
            public string Akronim { get; set; } = "";
        }
        private class Grafik
        {
            public Pracownik Pracownik { get; set; } = new();
            public int Miesiac { get; set; } = 0;
            public int Rok { get; set; } = 0;
            public List<Dane_Dnia> Dane_Dni { get; set; } = new();
            public string Nazwa_Pliku = "";
            public int Nr_Zakladki = 1;
            public void Set_Miesiac(string wartosc)
            {
                wartosc = wartosc.Trim().ToLower();
                if (wartosc.Contains("styczeń"))
                {
                    Miesiac = 1;
                }
                else if (wartosc.Contains("luty"))
                {
                    Miesiac = 2;
                }
                else if (wartosc.Contains("marzec"))
                {
                    Miesiac = 3;
                }
                else if (wartosc.Contains("kwiecień"))
                {
                    Miesiac = 4;
                }
                else if (wartosc.Contains("maj"))
                {
                    Miesiac = 5;
                }
                else if (wartosc.Contains("czerwiec"))
                {
                    Miesiac = 6;
                }
                else if (wartosc.Contains("lipiec"))
                {
                    Miesiac = 7;
                }
                else if (wartosc.Contains("sierpień"))
                {
                    Miesiac = 8;
                }
                else if (wartosc.Contains("wrzesień"))
                {
                    Miesiac = 9;
                }
                else if (wartosc.Contains("październik"))
                {
                    Miesiac = 10;
                }
                else if (wartosc.Contains("listopad"))
                {
                    Miesiac = 11;
                }
                else if (wartosc.Contains("grudzień"))
                {
                    Miesiac = 12;
                }
                else
                {
                    Miesiac = 0;
                }
            }
        }
        private class Dane_Dnia
        {
            public int Nr_Dnia { get; set; } = 0;
            public TimeSpan Godzina_Pracy_Od { get; set; } = TimeSpan.Zero;
            public TimeSpan Godzina_Pracy_Do { get; set; } = TimeSpan.Zero;
        }
        private class Current_Position
        {
            public int row = 1, col = 1;
        }
        private static List<Current_Position> Find_Grafiki(IXLWorksheet worksheet)
        {
            List<Current_Position> Lista_Pozycji_startowych_grafików = [];
            int Limiter = 1000;
            int counter = 0;
            foreach (var cell in worksheet.CellsUsed())
            {
                try
                {
                    if (cell.HasFormula && !cell.Address.ToString()!.Equals(cell.FormulaA1))
                    {
                        counter++;
                        if (counter > Limiter)
                        {
                            break;
                        }
                        continue;
                    }
                    if (cell.Value.ToString().Contains("Data"))
                    {
                        Lista_Pozycji_startowych_grafików.Add(new Current_Position()
                        {
                            row = cell.Address.RowNumber,
                            col = cell.Address.ColumnNumber
                        });
                    }
                }
                catch
                {
                    continue;
                }
            }
            return Lista_Pozycji_startowych_grafików;
        }
        private static void Process_Zakladka_For_Optima(IXLWorksheet worksheet)
        {
            var Lista_Pozycji_Grafików_Z_Zakladki = Find_Grafiki(worksheet);
            List<Grafik> grafiki = new();
            foreach (var Startpozycja in Lista_Pozycji_Grafików_Z_Zakladki)
            {
                var pozycja = Startpozycja;
                int counter = 0;
                while (true)
                {
                    Grafik grafik = new();
                    grafik.Pracownik = Get_Pracownik(worksheet, new Current_Position{ row = Startpozycja.row - 2, col = Startpozycja.col + ((counter * 3) + 1) });
                    if (string.IsNullOrEmpty(grafik.Pracownik.Imie) && string.IsNullOrEmpty(grafik.Pracownik.Nazwisko) && string.IsNullOrEmpty(grafik.Pracownik.Akronim))
                    {
                        break;
                    }
                    grafik.Dane_Dni = Get_Dane_Dni(worksheet, new Current_Position { row = Startpozycja.row + 4, col = Startpozycja.col + ((counter * 3) + 1) });
                    foreach( var g in grafik.Dane_Dni)
                    {
                        Console.WriteLine(g.Godzina_Pracy_Od);
                    }
                    grafiki.Add(grafik);
                    counter++;
                }
            }
        }
        private static Pracownik Get_Pracownik(IXLWorksheet worksheet, Current_Position pozycja)
        {
            Pracownik pracownik = new Pracownik();
            var nazwiskoimie = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
            if (!string.IsNullOrEmpty(nazwiskoimie))
            {
                if (nazwiskoimie.Split(" ").Length <= 1)
                {
                    //Program.error_logger.New_Error(dane, "Nazwisko i Imie", pozycja.col, pozycja.row, "Zły format wpisanego imienia i nazwiska pracownika");
                    // to gud zostaw to error zrob jesli akronim zly
                }
            }

            var akronim = "";
            for (int i = 0; i < 3; i++)
            {
                if (string.IsNullOrEmpty(akronim))
                {
                    akronim = worksheet.Cell(pozycja.row, pozycja.col + i).GetFormattedString().Trim();
                }
            }
            pracownik.Akronim = akronim;
            if (!string.IsNullOrEmpty(nazwiskoimie))
            {
                pracownik.Nazwisko = nazwiskoimie.Split(" ")[0];
                pracownik.Imie = nazwiskoimie.Split(" ")[1];
            }
            return pracownik;
        }
        private static List<Dane_Dnia> Get_Dane_Dni(IXLWorksheet worksheet, Current_Position pozycja)
        {
            List<Dane_Dnia> Dane_Dni = new();
            for(int i = 0; i < 31; i++){
                Dane_Dnia dane_Dnia = new Dane_Dnia();
                dane_Dnia.Nr_Dnia = i+1;
                var dane = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    pozycja.row += 1;
                    continue;
                }
                dane_Dnia.Godzina_Pracy_Od = TimeSpan.Parse(dane);
                dane = worksheet.Cell(pozycja.row, pozycja.col+1).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    pozycja.row += 1;
                    continue;
                }
                dane_Dnia.Godzina_Pracy_Do = TimeSpan.Parse(dane);
                Dane_Dni.Add(dane_Dnia);
                pozycja.row += 1;
            }
            return Dane_Dni;
        }






        public static int Main()
        {
            string filePath = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Kopia pliku WZÓR - GRAFIK (INDYWIDUALNY ROZKŁAD CZASU PRACY PRACOWNIKÓW).xlsx";
            int ilosc_zakladek = 0;
            using (var workbook = new XLWorkbook(filePath))
            {
                ilosc_zakladek = workbook.Worksheets.Count;
                for (int i = 1; i <= ilosc_zakladek; i++)
                {
                    var zakladka = workbook.Worksheet(i);
                    Process_Zakladka_For_Optima(zakladka);
                }
            }
            return 0;
        }
    }
}