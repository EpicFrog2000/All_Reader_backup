using ClosedXML.Excel;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.SqlClient;
using Microsoft.IdentityModel.Tokens;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace All_Readeer
{
    internal static class Karta_Pracy_Reader_v2
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
            public int Akronim { get; set; } = -1;
        }
        private class Karta_Pracy
        {
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
            public Pracownik pracownik { get; set; } = new();
            public int rok { get; set; } = 0;
            public int miesiac { get; set; } = 0;
            public int Set_Data(string Wartosc)
            {
                if (string.IsNullOrEmpty(Wartosc))
                {
                    return 1;
                }
                try
                {
                    DateTime data;
                    if (!DateTime.TryParse(Wartosc, out data))
                    {
                        return 1;
                    }
                    rok = data.Year;
                    miesiac = data.Month;
                } catch {
                    return 1;
                }
                return 0;
            }
            public List<Dane_Dnia> dane_dni { get; set; } = [];
            public List<Nieobecnosc> ListaNieobecnosci { get; set; } = [];
            public void Set_Miesiac(string nazwa)
            {
                if (!string.IsNullOrEmpty(nazwa))
                {
                    if (nazwa.ToLower() == "styczeń")
                    {
                        miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "i")
                    {
                        miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "luty")
                    {
                        miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "ii")
                    {
                        miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "marzec")
                    {
                        miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "iii")
                    {
                        miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "kwiecień")
                    {
                        miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "iv")
                    {
                        miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "maj")
                    {
                        miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "v")
                    {
                        miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "czerwiec")
                    {
                        miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "vi")
                    {
                        miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "lipiec")
                    {
                        miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "vii")
                    {
                        miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "sierpień")
                    {
                        miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "viii")
                    {
                        miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "wrzesień")
                    {
                        miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "ix")
                    {
                        miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "październik")
                    {
                        miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "x")
                    {
                        miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "listopad")
                    {
                        miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "xi")
                    {
                        miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "grudzień")
                    {
                        miesiac = 12;
                    }
                    else if (nazwa.ToLower() == "xii")
                    {
                        miesiac = 12;
                    }
                    else
                    {
                        miesiac = 0;
                    }
                }
            }
        }
        private class Dane_Dnia
        {
            public int dzien { get; set; } = 0;
            public TimeSpan godz_rozp_pracy { get; set; } = TimeSpan.Zero;
            public TimeSpan godz_zakoncz_pracy { get; set; } = TimeSpan.Zero;
            public decimal praca_wg_grafiku { get; set; } = 0;
            public decimal liczba_godz_przepracowanych { get; set; } = 0;
            public decimal Godz_nadl_platne_z_dod_50 { get; set; } = 0;
            public decimal Godz_nadl_platne_z_dod_100 { get; set; } = 0;
            public decimal Godz_Odbior { get; set; } = 0;
            public DateTime Dzien_Odbior { get; set; } = DateTime.MinValue;
        }
        private class Current_Position
        {
            public int row { get; set; } = 1;
            public int col { get; set; } = 1;
        }
        private enum RodzajNieobecnosci
        {
            DE,     // Delegacja
            DM,     // Dodatkowy urlop macierzyński
            DR,     // Urlop rodzicielski
            IK,     // Izolacja - Koronawirus
            NB,     // Badania lekarskie - okresowe
            NN,     // Nieobecność nieusprawiedliwiona
            NR,     // Badania lekarskie - z tyt. niepełnosprawności
            NU,     // Nieobecność usprawiedliwiona
            OD,     // Oddelegowanie do prac w ZZ
            OG,     // Odbiór godzin dyżuru
            ON,     // Odbiór nadgodzin
            OO,     // Odbiór pracy w niedziele
            OP,     // Urlop opiekuńczy (niepłatny)
            OS,     // Odbiór pracujących sobót
            PP,     // Poszukiwanie pracy
            PZ,     // Praca zdalna okazjonalna
            SW,     // Urlop/zwolnienie z tyt. siły wyższej
            SZ,     // Szkolenie
            SP,     // Zwolniony z obowiązku świadcz. pracy
            U9,     // Urlop rodzicielski 9 tygodni
            UA,     // Długotrwały urlop bezpłatny
            UB,     // Urlop bezpłatny
            UC,     // Urlop ojcowski
            UD,     // Na opiekę nad dzieckiem art.K.P.188
            UJ,     // Ćwiczenia wojskowe
            UK,     // Urlop dla krwiodawcy
            UL,     // Służba wojskowa
            ULawnika, // Praca ławnika w sądzie
            UM,     // Urlop macierzyński
            UN,     // Urlop z tyt. niepełnosprawności
            UO,     // Urlop okolicznościowy
            UP,     // Dodatkowy urlop osoby represjonowanej
            UR,     // Dodatkowe dni na turnus rehabilitacyjny
            US,     // Urlop szkoleniowy
            UV,     // Urlop weterana
            UW,     // Urlop wypoczynkowy
            UY,     // Urlop wychowawczy
            UZ,     // Urlop na żądanie
            WY,     // Wypoczynek skazanego
            ZC,     // Opieka nad członkiem rodziny (ZLA)
            ZD,     // Opieka nad dzieckiem (ZUS ZLA)
            ZK,     // Opieka nad dzieckiem Koronawirus
            ZL,     // Zwolnienie lekarskie (ZUS ZLA)
            ZN,     // Zwolnienie lekarskie niepłatne (ZLA)
            ZP,     // Kwarantanna sanepid
            ZR,     // Zwolnienie na rehabilitację (ZUS ZLA)
            ZS,     // Zwolnienie szpitalne (ZUS ZLA)
            ZY,     // Zwolnienie powypadkowe (ZUS ZLA)
            ZZ      // Zwolnienie lek. (ciąża) (ZUS ZLA)
        }
        private class Nieobecnosc
        {
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
            public Pracownik pracownik = new();
            public int rok = 0;
            public int miesiac = 0;
            public int dzien = 0;
            public RodzajNieobecnosci rodzaj_absencji = 0;
        }
        public static void Process_Zakladka_For_Optima(IXLWorksheet worksheet)
        {
            try
            {
                List<Karta_Pracy> karty_pracy = [];
                var tabelki = Find_Karty(worksheet);
                foreach(var tabelka in tabelki)
                {
                    Karta_Pracy karta_pracy = new();
                    karta_pracy.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                    karta_pracy.nr_zakladki = Program.error_logger.Nr_Zakladki;
                    Get_Header_Karta_Info(tabelka, worksheet, ref karta_pracy);
                    Get_Dane_Dni(tabelka, worksheet, ref karta_pracy);
                    karty_pracy.Add(karta_pracy);
                }
                if (karty_pracy.Count > 0)
                {
                    foreach (var karta in karty_pracy)
                    {
                        try
                        {
                            Dodaj_Dane_Do_Optimy(karta);
                        }
                        catch
                        {
                            throw;
                        }
                    }
                }
            }
            catch(Exception ex){
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        private static List<Current_Position> Find_Karty(IXLWorksheet worksheet)
        {
            List<Current_Position> starty = new();
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
                    if (cell.Value.ToString().Contains("Dzień"))
                    {
                        starty.Add(new Current_Position()
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
            return starty;
        }
        private static void Get_Header_Karta_Info(Current_Position StartKarty, IXLWorksheet worksheet, ref Karta_Pracy karta_pracy)
        {
            //wczytaj date
            var dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + 4).GetFormattedString().Trim().ToLower();
            for (int i = 0; i < 12; i++)
            {
                if (string.IsNullOrEmpty(dane))
                {
                    dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + 4 + i).GetFormattedString().Trim().ToLower();
                }
                else
                {
                    //here try to get data i rok
                    if (dane.EndsWith("r"))
                    {
                        dane = dane.Substring(0, dane.Length - 1).Trim();
                    }
                    if (dane.EndsWith("r."))
                    {
                        dane = dane.Substring(0, dane.Length - 2).Trim();
                    }

                    string[] dateFormats = { "dd.MM.yyyy" };
                    if (DateTime.TryParseExact(dane, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedData))
                    {
                        karta_pracy.miesiac = parsedData.Month;
                        karta_pracy.rok = parsedData.Year;
                    }
                    else
                    {
                        if (dane.Contains("pażdziernik"))
                        {
                            dane = dane.Replace("pażdziernik", "październik");
                        }
                        if (karta_pracy.Set_Data(dane) == 1)
                        {
                            if (dane.Split(" ").Length == 2)
                            {
                                var ndata = dane.Split(" ");
                                try
                                {
                                    karta_pracy.Set_Miesiac(ndata[0]);
                                    if (int.TryParse(Regex.Replace(ndata[1], @"\D", ""), out int rok))
                                    {
                                        karta_pracy.rok = rok;
                                    }
                                }
                                catch { }
                            }
                            else if (dane.Split(" ").Length == 3)
                            {
                                var ndata = dane.Split(" ");
                                try
                                {
                                    karta_pracy.Set_Miesiac(ndata[1]);
                                    if (int.TryParse(ndata[2], out int rok))
                                    {
                                        karta_pracy.rok = rok;
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                if (dane.Split(" ").Count() > 1)
                                {
                                    //wez 2 od tylu
                                    var ndata = dane.Split(" ");
                                    try
                                    {
                                        karta_pracy.Set_Miesiac(ndata[^2]);
                                        if (int.TryParse(ndata[^1], out int rok))
                                        {
                                            karta_pracy.rok = rok;
                                        }
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    if (karta_pracy.miesiac == 0 || karta_pracy.rok == 0)
                    {
                        dane = worksheet.Cell(StartKarty.row - 4, StartKarty.col + 4 + i - 1).GetFormattedString().Trim().ToLower();
                        if (!string.IsNullOrEmpty(dane) && DateTime.TryParseExact(dane, dateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedData2))
                        {
                            karta_pracy.miesiac = parsedData2.Month;
                            karta_pracy.rok = parsedData2.Year;
                        }
                    }

                    if (karta_pracy.miesiac != 0 && karta_pracy.rok != 0)
                    {
                        break;
                    }
                }
            }
            if (karta_pracy.miesiac == 0 || karta_pracy.rok == 0)
            {
                Program.error_logger.New_Error(dane, "data", StartKarty.col + 11, StartKarty.row - 3, $"Nie wykryto daty w pliku. Oczekiwana dat między kolumna[{StartKarty.col + 4}] rząd[{StartKarty.row - 3}] a kolumna[{StartKarty.col + 4 + 11}] rząd[{StartKarty.row - 3}]");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            //wczytaj nazwisko i imie
            try{
                Current_Position pozycja_wczytania_danych = new();
                int tmpi = 0;
                string[] wordsToRemove = { "IMIĘ:", "IMIE:", "NAZWISKO:", "NAZWISKO", " IMIE", "IMIĘ", ":" };
                dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col).GetFormattedString().Trim().Replace("  ", " ");
                pozycja_wczytania_danych.row = StartKarty.row - 2;
                pozycja_wczytania_danych.col = StartKarty.col;
                for (int i = 0; i < 6; i++)
                {
                    foreach (var word in wordsToRemove)
                    {
                        var pattern = $@"\b{Regex.Escape(word)}\b";
                        dane = Regex.Replace(dane, pattern, "", RegexOptions.IgnoreCase);
                    }

                    dane = Regex.Replace(dane, @"\s+", " ").Trim();
                    if (dane.Contains("KARTA PRACY:"))
                    {
                        dane = dane.Replace("KARTA PRACY:", "").Trim();
                    }
                    if (dane.Contains("KARTA PRACY"))
                    {
                        dane = dane.Replace("KARTA PRACY", "").Trim();
                    }
                    if (!string.IsNullOrEmpty(dane))
                    {
                        tmpi = i;
                        break;
                    }
                    else
                    {
                        dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col + i).GetFormattedString().Trim().Replace("  ", " ");
                        pozycja_wczytania_danych.row = StartKarty.row - 2;
                        pozycja_wczytania_danych.col = StartKarty.col + i;
                    }
                }
                if (string.IsNullOrEmpty(dane))
                {
                    dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col).GetFormattedString().Trim().Replace("  ", " ");
                    pozycja_wczytania_danych.row = StartKarty.row - 3;
                    pozycja_wczytania_danych.col = StartKarty.col;
                    for (int i = 0; i < 6; i++)
                    {
                        foreach (var word in wordsToRemove)
                        {
                            dane = dane.Replace(word, "", StringComparison.OrdinalIgnoreCase);
                        }
                        dane = Regex.Replace(dane, @"\s+", " ").Trim();
                        if (dane.Contains("KARTA PRACY:"))
                        {
                            dane = dane.Replace("KARTA PRACY:", "").Trim();
                        }
                        if (!string.IsNullOrEmpty(dane))
                        {
                            tmpi = i;
                            break;
                        }
                        else
                        {
                            dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + i).GetFormattedString().Trim().Replace("  ", " ");
                            pozycja_wczytania_danych.row = StartKarty.row - 3;
                            pozycja_wczytania_danych.col = StartKarty.col + i;
                        }
                    }
                }
                foreach (var word in wordsToRemove)
                {
                    dane = dane.Replace(word, "", StringComparison.OrdinalIgnoreCase);
                }
                dane = Regex.Replace(dane, @"\s+", " ").Trim();
                if (dane.Contains("KARTA PRACY:"))
                {
                    dane = dane.Replace("KARTA PRACY:", "").Trim();
                }
                if (string.IsNullOrEmpty(dane))
                {
                    //Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, $"Nie znaleziono pola z nazwiskiem i imieniem między kolumna[{StartKarty.col}] rząd[{StartKarty.row - 2}] a kolumna[{StartKarty.col + 5}] rząd[{StartKarty.row - 2}]");
                    //throw new Exception(Program.error_logger.Get_Error_String());
                }
                else
                {
                    try
                    {
                        var parts = dane.Trim().Split(' ');
                        if (parts.Length == 2)
                        {
                            karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                            karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[1];
                        }else if (parts.Length == 3)
                        {
                            if (int.TryParse(parts[0], out int parsedValue))
                            {
                                karta_pracy.pracownik.Akronim = parsedValue;
                                karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[1];
                                karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[2];
                            }else if (int.TryParse(parts[2], out int parsedValue2))
                            {
                                karta_pracy.pracownik.Akronim = parsedValue2;
                                karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                                karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[1];
                            }
                        }
                    }
                    catch
                    {
                        Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Zły format pola nazwisko i imie. Powinno być: KARTA PRACY: Nazwisko Imie");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }

                    if (string.IsNullOrEmpty(karta_pracy.pracownik.Imie) && !string.IsNullOrEmpty(karta_pracy.pracownik.Nazwisko))
                    {
                        dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + tmpi).GetFormattedString().Trim().Replace("  ", " ");
                        pozycja_wczytania_danych.row = StartKarty.row - 3;
                        pozycja_wczytania_danych.col = StartKarty.col + tmpi;
                        if (!string.IsNullOrEmpty(dane))
                        {
                            karta_pracy.pracownik.Imie = dane;
                        }
                    }
                    else if (string.IsNullOrEmpty(karta_pracy.pracownik.Nazwisko) && !string.IsNullOrEmpty(karta_pracy.pracownik.Imie))
                    {
                        dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + tmpi).GetFormattedString().Trim().Replace("  ", " ");
                        pozycja_wczytania_danych.row = StartKarty.row - 3;
                        pozycja_wczytania_danych.col = StartKarty.col + tmpi;
                        if (!string.IsNullOrEmpty(dane))
                        {
                            karta_pracy.pracownik.Nazwisko = dane;
                        }
                    }
                    if (string.IsNullOrEmpty(karta_pracy.pracownik.Imie) || string.IsNullOrEmpty(karta_pracy.pracownik.Nazwisko))
                    {
                        //Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, $"Nie znaleziono pola z nazwiskiem i imieniem między kolumna[{StartKarty.col}] rząd[{StartKarty.row - 2}] a kolumna[{StartKarty.col + 5}] rząd[{StartKarty.row - 2}]");
                        //throw new Exception(Program.error_logger.Get_Error_String());
                    }

                }

                // znajdz akronim w prawo
                if (karta_pracy.pracownik.Akronim == -1)
                {
                    var akronim = "";
                    for (int i = 0; i < 6; i++)
                    {
                        akronim = worksheet.Cell(pozycja_wczytania_danych.row, pozycja_wczytania_danych.col + 1 + i).GetFormattedString().Trim().Replace("  ", " ");
                        if (!string.IsNullOrEmpty(akronim) && Regex.IsMatch(akronim, @"^(akronim|akronim:\s*\d*|\d+)$"))
                        {
                            break;
                        }
                    }
                    if (!string.IsNullOrEmpty(akronim))
                    {
                        if (int.TryParse(akronim, out int parsedValue))
                        {
                            karta_pracy.pracownik.Akronim = parsedValue;
                        }
                        else
                        {
                            try
                            {
                                if (int.TryParse(akronim.Split(' ')[1], out int parsedValue2))
                                {
                                    karta_pracy.pracownik.Akronim = parsedValue2;
                                }
                            }
                            catch
                            {
                                karta_pracy.pracownik.Akronim = -1;
                            }
                            karta_pracy.pracownik.Akronim = -1;
                        }
                    }
                    // Jeśli nie znalazlo akronim, to możę jest obok imie nazwisko w tej samej komórce
                    if(karta_pracy.pracownik.Akronim == -1)
                    {
                        if (int.TryParse(dane.Trim().Split(' ')[^1], out int parsedValue3))
                        {
                            karta_pracy.pracownik.Akronim = parsedValue3;
                        }
                    }
                }

                if ((karta_pracy.pracownik.Nazwisko == null || karta_pracy.pracownik.Imie == null) || (string.IsNullOrEmpty(karta_pracy.pracownik.Nazwisko) || string.IsNullOrEmpty(karta_pracy.pracownik.Imie)))
                {
                    if (karta_pracy.pracownik.Akronim == 0)
                    {
                        Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Zły format pola nazwisko i imie. Powinno być: KARTA PRACY: Nazwisko Imie");
                        Program.error_logger.New_Error(dane, "akronim", pozycja_wczytania_danych.col + 2, pozycja_wczytania_danych.row, "Nie znaleziono wartości akronim. Powinno być: Akronim");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
                if(karta_pracy.pracownik.Nazwisko != null && !string.IsNullOrEmpty(karta_pracy.pracownik.Nazwisko))
                {
                    karta_pracy.pracownik.Nazwisko = karta_pracy.pracownik.Nazwisko.ToLower();
                    karta_pracy.pracownik.Nazwisko = char.ToUpper(karta_pracy.pracownik.Nazwisko[0], CultureInfo.CurrentCulture) + karta_pracy.pracownik.Nazwisko.Substring(1);
                }
                if (karta_pracy.pracownik.Imie != null && string.IsNullOrEmpty(karta_pracy.pracownik.Imie))
                {
                    karta_pracy.pracownik.Imie = karta_pracy.pracownik.Imie.ToLower();
                    karta_pracy.pracownik.Imie = char.ToUpper(karta_pracy.pracownik.Imie[0], CultureInfo.CurrentCulture) + karta_pracy.pracownik.Imie.Substring(1);
                }

            }
            catch (Exception ex)
            {
                Program.error_logger.New_Error(dane, "Imie nazwisko", StartKarty.row - 2, StartKarty.col, "Nieznany format");
                throw new Exception($"{Program.error_logger.Get_Error_String()}");
            }
        }
        private static void Get_Dane_Dni(Current_Position StartKarty, IXLWorksheet worksheet, ref Karta_Pracy karta_pracy)
        {
            StartKarty.row += 3;
            var NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetFormattedString().Trim();
            while (!string.IsNullOrEmpty(NrDnia))
            {
                Dane_Dnia dzien = new();
                // dzien miesiaca
                if (int.TryParse(NrDnia, out int parsedDzien))
                {
                    dzien.dzien = parsedDzien;
                }else if (DateTime.TryParse(NrDnia, out DateTime Data))
                {
                    dzien.dzien = Data.Day;
                }else
                {
                    Program.error_logger.New_Error(NrDnia, "dzien", StartKarty.col, StartKarty.row, "Błędny nr dnia");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                var Cell_Value = "";

                // try get odb nadgodzin
                try
                {
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 5).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(Cell_Value))
                    {
                        if (decimal.TryParse(Cell_Value, out decimal ilosc))
                        {
                            if (ilosc != 0)
                            {
                                dzien.Godz_Odbior = ilosc;
                                karta_pracy.dane_dni.Add(dzien);
                            }
                        }
                        else
                        {
                            Program.error_logger.New_Error(Cell_Value, "liczba godzin odbioru za prace w nadgodz", StartKarty.col + 5, StartKarty.row, "Niepoprawnie wpisana liczba");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                }
                catch
                {
                    throw;
                }

                //try get nieobecność:
                try
                {
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 3).GetFormattedString().Trim();
                    if (!string.IsNullOrEmpty(Cell_Value) && Cell_Value != "PZ")
                    {
                        Nieobecnosc nieobecnosc = new();
                        if (RodzajNieobecnosci.TryParse(Cell_Value.ToUpper(), out RodzajNieobecnosci Rnieobecnosc))
                        {
                            nieobecnosc.rodzaj_absencji = Rnieobecnosc;
                            nieobecnosc.pracownik = karta_pracy.pracownik;
                            nieobecnosc.rok = karta_pracy.rok;
                            nieobecnosc.miesiac = karta_pracy.miesiac;
                            nieobecnosc.dzien = dzien.dzien;
                            nieobecnosc.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                            nieobecnosc.nr_zakladki = Program.error_logger.Nr_Zakladki;
                        }
                        else
                        {
                            Program.error_logger.New_Error(Cell_Value, "Kod absencji", StartKarty.col + 3, StartKarty.row, "Nieprawidłowy kod nieobecności");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        karta_pracy.ListaNieobecnosci.Add(nieobecnosc);
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetFormattedString().Trim();
                        continue;
                    }
                }
                catch
                {
                    throw;
                }

                // godz rozpoczecia
                try
                {
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 1).GetFormattedString().Trim();
                    if (!string.IsNullOrEmpty(Cell_Value))
                    {
                        dzien.godz_rozp_pracy = Reader.Try_Get_Date(Cell_Value);
                    }
                }
                catch
                {
                    Program.error_logger.New_Error(Cell_Value, "Godzina Rozpoczęcia Pracy", StartKarty.col + 1, StartKarty.row, "Powinna byc godzina w formacie np. 08:00");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // godz zakonczenia
                try
                {
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 2).GetFormattedString().Trim();
                    if (!string.IsNullOrEmpty(Cell_Value))
                    {
                        dzien.godz_zakoncz_pracy = Reader.Try_Get_Date(Cell_Value);
                    }
                }
                catch
                {
                    Program.error_logger.New_Error(Cell_Value, "Godzina Zakończenia Pracy", StartKarty.col + 1, StartKarty.row, "Powinna byc godzina w formacie np. 08:00");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                //get godz_nad 50
                Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 9).GetFormattedString().Trim();
                if (!string.IsNullOrEmpty(Cell_Value))
                {
                    if (decimal.TryParse(Cell_Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var godzNadlPlatne))
                    {
                        dzien.Godz_nadl_platne_z_dod_50 = godzNadlPlatne;
                    }
                    else
                    {
                        Program.error_logger.New_Error(Cell_Value, "Liczba godzin z dodatkiem 50%", StartKarty.col + 1, StartKarty.row, "Niepoprawny format w polu");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
                //get godz_nad 100
                Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 10).GetFormattedString().Trim();
                if (!string.IsNullOrEmpty(Cell_Value))
                {
                    if (decimal.TryParse(Cell_Value, NumberStyles.Any, CultureInfo.InvariantCulture, out var godzNadlPlatne))
                    {
                        dzien.Godz_nadl_platne_z_dod_100 = godzNadlPlatne;
                    }
                    else
                    {
                        Program.error_logger.New_Error(Cell_Value, "Liczba godzin z dodatkiem 100%", StartKarty.col + 1, StartKarty.row, "Niepoprawny format w polu");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }

                if (dzien.godz_rozp_pracy != TimeSpan.Zero && dzien.godz_zakoncz_pracy != TimeSpan.Zero)
                {
                    karta_pracy.dane_dni.Add(dzien);
                }
                StartKarty.row++;
                NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetFormattedString().Trim();
            }
        }
        private static int Dodaj_Obecnosci_do_Optimy(Karta_Pracy karta, SqlTransaction tran, SqlConnection connection)
        {
            
            int ilosc_wpisow = 0;
            DateTime baseDate = new DateTime(1899, 12, 30);
            foreach (var dane_Dni in karta.dane_dni)
            {
                try
                {
                    if (DateTime.TryParse($"{karta.rok}-{karta.miesiac:D2}-{dane_Dni.dzien:D2}", out DateTime Data_Karty))
                    {
                        var (startPodstawowy, endPodstawowy, startNadl50, endNadl50, startNadl100, endNadl100) = Oblicz_Czas_Z_Dodatkiem(dane_Dni);
                        double czasPrzepracowany = 0;
                        if (dane_Dni.godz_zakoncz_pracy < dane_Dni.godz_rozp_pracy)
                        {
                            czasPrzepracowany = (TimeSpan.FromHours(24) - dane_Dni.godz_rozp_pracy).TotalHours + dane_Dni.godz_zakoncz_pracy.TotalHours;
                        }
                        else
                        {
                            czasPrzepracowany = (dane_Dni.godz_zakoncz_pracy - dane_Dni.godz_rozp_pracy).TotalHours;
                        }
                        double czasPodstawowy = czasPrzepracowany - ((double)(dane_Dni.Godz_nadl_platne_z_dod_50 + dane_Dni.Godz_nadl_platne_z_dod_100));
                        if (czasPodstawowy > 0)
                        {
                            ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, tran, Data_Karty, startPodstawowy, endPodstawowy, karta, 2); // 2 = czas PP
                        }
                        if (dane_Dni.Godz_nadl_platne_z_dod_50 > 0)
                        {
                            ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, tran, Data_Karty, startNadl50, endNadl50, karta, 2);
                        }
                        if (dane_Dni.Godz_nadl_platne_z_dod_100 > 0)
                        {
                            ilosc_wpisow += Zrob_Insert_Obecnosc_Command(connection, tran, Data_Karty, startNadl100, endNadl100, karta, 2);
                        }
                    }
                }
                catch (SqlException ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
                catch (FormatException)
                {
                    continue;
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
                
            }
            return ilosc_wpisow;
        }                                                                                                                                                                                    //  2 -> podstawowy
        private static int Zrob_Insert_Obecnosc_Command(SqlConnection connection, SqlTransaction transaction, DateTime Data_Karty, TimeSpan startPodstawowy, TimeSpan endPodstawowy, Karta_Pracy karta, int Typ_Pracy)
        {
            DateTime baseDate = new DateTime(1899, 12, 30);
            DateTime godzOdDate = baseDate + startPodstawowy;
            DateTime godzDoDate = baseDate + endPodstawowy;
            bool duplicate = false;
            using (SqlCommand cmd = new SqlCommand(@"
    IF EXISTS (
        SELECT 1
        FROM cdn.PracPracaDni P
        INNER JOIN CDN.PracPracaDniGodz G ON P.PPR_PprId = G.PGR_PprId
        WHERE P.PPR_PraId = @PRI_PraId 
          AND P.PPR_Data = @DataInsert
          AND G.PGR_OdGodziny = @GodzOdDate
          AND G.PGR_DoGodziny = @GodzDoDate
          AND G.PGR_Strefa = @TypPracy
    )
    BEGIN
        SELECT 1;
    END
    ELSE
    BEGIN
        SELECT 0;
    END", connection, transaction))
            {
                cmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                cmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                cmd.Parameters.AddWithValue("@DataInsert", Data_Karty);
                cmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(karta.pracownik));
                cmd.Parameters.AddWithValue("@TypPracy", Typ_Pracy);
                int result = (int)cmd.ExecuteScalar();
                duplicate = (result == 1);
            }

            if (!duplicate)
            {
                if (godzOdDate != godzDoDate)
                {
                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, transaction))
                    {
                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        insertCmd.Parameters.AddWithValue("@DataInsert", Data_Karty);
                        insertCmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(karta.pracownik));
                        insertCmd.Parameters.AddWithValue("@TypPracy", Typ_Pracy);
                        if (Program.error_logger.Last_Mod_Osoba.Length > 20)
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 20));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba);
                        }
                        if (Program.error_logger.Last_Mod_Osoba.Length > 50)
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 50));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba);
                        }
                        insertCmd.Parameters.AddWithValue("@DataMod", Program.error_logger.Last_Mod_Time);
                        insertCmd.ExecuteScalar();
                    }
                }
                return 1;
            }
            return 0;
        }
        private static int Dodaj_Nieobecnosci_do_Optimy(List<Nieobecnosc> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
        {
            int ilosc_wpisow = 0;
            DateTime baseDate = new DateTime(1899, 12, 30);
            List<List<Nieobecnosc>> Nieobecnosci = Podziel_Niobecnosci_Na_Osobne(ListaNieobecności);
            foreach (var ListaNieo in Nieobecnosci)
            {
                DateTime dataniobecnoscistart;
                DateTime dataniobecnosciend;
                try
                {
                    dataniobecnoscistart = new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien);
                    dataniobecnosciend = new DateTime(ListaNieo[ListaNieo.Count - 1].rok, ListaNieo[ListaNieo.Count - 1].miesiac, ListaNieo[ListaNieo.Count - 1].dzien);
                }
                catch
                {
                    continue;
                }

                int przyczyna = Dopasuj_Przyczyne(ListaNieo[0].rodzaj_absencji);
                var nazwa_nieobecnosci = Dopasuj_Nieobecnosc(ListaNieo[0].rodzaj_absencji);
                if (string.IsNullOrEmpty(nazwa_nieobecnosci))
                {
                    Program.error_logger.New_Custom_Error($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                    var e = new Exception($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                    e.Data["Kod"] = 42069;
                    throw e;
                }
                var dni_robocze = Ile_Dni_Roboczych(ListaNieo);
                var dni_calosc = ListaNieo.Count;

                bool duplicate = false;

                using (SqlCommand cmd = new SqlCommand(@"IF EXISTS (
SELECT 1 
FROM CDN.PracNieobec
WHERE [PNB_PraId] = @PRI_PraId
    AND [PNB_TnbId] = (
        SELECT TNB_TnbId 
        FROM cdn.TypNieobec 
        WHERE TNB_Nazwa = @NazwaNieobecnosci
    )
    AND [PNB_OkresOd] = @DataOd
    AND [PNB_OkresDo] = @DataDo
    AND [PNB_RozliczData] = @BaseDate
    AND [PNB_Przyczyna] = @Przyczyna
    AND [PNB_DniPracy] = @DniPracy
    AND [PNB_DniKalend] = @DniKalendarzowe
)
BEGIN
SELECT 1
END
ELSE 
BEGIN
SELECT 0
END
", connection, tran))
                {
                    cmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(ListaNieo[0].pracownik));
                    cmd.Parameters.AddWithValue("@NazwaNieobecnosci", nazwa_nieobecnosci);
                    cmd.Parameters.AddWithValue("@DniPracy", dni_robocze);
                    cmd.Parameters.AddWithValue("@DniKalendarzowe", dni_calosc);
                    cmd.Parameters.AddWithValue("@Przyczyna", przyczyna);
                    cmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = dataniobecnoscistart;
                    cmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = baseDate;
                    cmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = dataniobecnosciend;
                    if ((int)cmd.ExecuteScalar() == 1)
                    {
                        duplicate = true;
                        return 0;
                    }
                }
                
                if (!duplicate)
                {
                    try
                    {
                        using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertNieObecnoŚciDoOptimy, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(ListaNieo[0].pracownik));
                            insertCmd.Parameters.AddWithValue("@NazwaNieobecnosci", nazwa_nieobecnosci);
                            insertCmd.Parameters.AddWithValue("@DniPracy", dni_robocze);
                            insertCmd.Parameters.AddWithValue("@DniKalendarzowe", dni_calosc);
                            insertCmd.Parameters.AddWithValue("@Przyczyna", przyczyna);
                            insertCmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = dataniobecnoscistart;
                            insertCmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = baseDate;
                            insertCmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = dataniobecnosciend;
                            if (Program.error_logger.Last_Mod_Osoba.Length > 20)
                            {
                                insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 20));
                            }
                            else
                            {
                                insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba);
                            }
                            if (Program.error_logger.Last_Mod_Osoba.Length > 50)
                            {
                                insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 50));
                            }
                            else
                            {
                                insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba);
                            }
                            insertCmd.Parameters.AddWithValue("@DataMod", Program.error_logger.Last_Mod_Time);
                            insertCmd.ExecuteScalar();
                        }
                    }
                    catch (SqlException ex)
                    {
                        tran.Rollback();
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    }
                    catch (FormatException)
                    {
                        continue;
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        if (ex.Data.Contains("kod") && ex.Data["kod"] is int kod && kod == 42069)
                        {
                            throw;
                        }
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    }
                    ilosc_wpisow++;
                }
            }
            return ilosc_wpisow;
        }
        private static string Dopasuj_Nieobecnosc(RodzajNieobecnosci rodzaj)
        {
            return rodzaj switch
            {
                RodzajNieobecnosci.UO => "Urlop okolicznościowy",
                RodzajNieobecnosci.ZL => "Zwolnienie chorobowe/F",
                RodzajNieobecnosci.ZY => "Zwolnienie chorobowe/wyp.w drodze/F",
                RodzajNieobecnosci.ZS => "Zwolnienie chorobowe/wyp.przy pracy/F",
                RodzajNieobecnosci.ZN => "Zwolnienie chorobowe/bez prawa do zas.",
                RodzajNieobecnosci.ZP => "Zwolnienie chorobowe/pozbawiony prawa",
                RodzajNieobecnosci.UR => "Urlop rehabilitacyjny",
                RodzajNieobecnosci.ZR => "Urlop rehabilitacyjny/wypadek w drodze..",
                RodzajNieobecnosci.ZD => "Urlop rehabilitacyjny/wypadek przy pracy",
                RodzajNieobecnosci.UM => "Urlop macierzyński",
                RodzajNieobecnosci.UC => "Urlop ojcowski",
                RodzajNieobecnosci.OP => "Urlop opiekuńczy (zasiłek)",
                RodzajNieobecnosci.UY => "Urlop wychowawczy (121)",
                RodzajNieobecnosci.UW => "Urlop wypoczynkowy",
                RodzajNieobecnosci.NU => "Nieobecność usprawiedliwiona (151)",
                RodzajNieobecnosci.NN => "Nieobecność nieusprawiedliwiona (152)",
                RodzajNieobecnosci.UL => "Służba wojskowa",
                RodzajNieobecnosci.DR => "Urlop rodzicielski",
                RodzajNieobecnosci.DM => "Urlop macierzyński dodatkowy",
                RodzajNieobecnosci.PP => "Dni wolne na poszukiwanie pracy",
                RodzajNieobecnosci.UK => "Dni wolne z tyt. krwiodawstwa",
                RodzajNieobecnosci.IK => "Covid19",
                _ => "Nieobecność (B2B)"
            };
        }
        private static int Dopasuj_Przyczyne(RodzajNieobecnosci rodzaj)
        {
            return rodzaj switch
            {
                RodzajNieobecnosci.ZL => 1,        // Zwolnienie lekarskie
                RodzajNieobecnosci.DM => 2,        // Urlop macierzyński
                RodzajNieobecnosci.DR => 13,        // Urlop opiekuńczy
                RodzajNieobecnosci.NB => 1,        // Zwolnienie lekarskie
                RodzajNieobecnosci.NN => 5,        // Nieobecność nieusprawiedliwiona
                RodzajNieobecnosci.UC => 21,       // Urlop opiekuńczy
                RodzajNieobecnosci.UD => 21,       // Urlop opiekuńczy
                RodzajNieobecnosci.UJ => 10,       // Służba wojskowa
                RodzajNieobecnosci.UL => 10,       // Służba wojskowa
                RodzajNieobecnosci.UM => 2,       // Urlop macierzyński
                RodzajNieobecnosci.UO => 4,       // Urlop okolicznościowy
                RodzajNieobecnosci.UN => 3,       // Urlop rehabilitacyjny
                RodzajNieobecnosci.UR => 3,       // Urlop rehabilitacyjny
                RodzajNieobecnosci.ZC => 21,       // Urlop opiekuńczy
                RodzajNieobecnosci.ZD => 21,       // Urlop opiekuńczy
                RodzajNieobecnosci.ZK => 21,       // Urlop opiekuńczy
                RodzajNieobecnosci.ZN => 1,       // Zwolnienie lekarskie
                RodzajNieobecnosci.ZR => 3,       // Urlop rehabilitacyjny
                RodzajNieobecnosci.ZZ => 1,       // Zwolnienie lekarskie
                _ => 9                             // Nie dotyczy dla pozostałych przypadków
            };
        }
        private static List<List<Nieobecnosc>> Podziel_Niobecnosci_Na_Osobne(List<Nieobecnosc> listaNieobecnosci)
        {
            List<List<Nieobecnosc>> listaOsobnychNieobecnosci = new();
            List<Nieobecnosc> currentGroup = new();

            foreach (var nieobecnosc in listaNieobecnosci)
            {
                if (currentGroup.Count == 0 || nieobecnosc.dzien == currentGroup[^1].dzien + 1)
                {
                    currentGroup.Add(nieobecnosc);
                }
                else
                {
                    listaOsobnychNieobecnosci.Add(new List<Nieobecnosc>(currentGroup));
                    currentGroup = new List<Nieobecnosc> { nieobecnosc };
                }
            }

            if (currentGroup.Count > 0)
            {
                listaOsobnychNieobecnosci.Add(currentGroup);
            }

            return listaOsobnychNieobecnosci;
        }
        private static void Dodaj_Dane_Do_Optimy(Karta_Pracy karta)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
                {
                    connection.Open();
                    using (SqlTransaction tran = connection.BeginTransaction())
                    {
                        if (karta.dane_dni.Count > 0)
                        {
                            if (Dodaj_Godz_Odbior_Do_Optimy(karta, tran, connection) > 0)
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"Poprawnie dodawno odbiory nadgodzin z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                            if (Dodaj_Obecnosci_do_Optimy(karta, tran, connection) > 0)
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"Poprawnie dodawno obecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                        if (karta.ListaNieobecnosci.Count > 0)
                        {
                            if(Dodaj_Nieobecnosci_do_Optimy(karta.ListaNieobecnosci, tran, connection) > 0)
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine($"Poprawnie dodawno nieobecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                                Console.ForegroundColor = ConsoleColor.White;
                            }
                        }
                        tran.Commit();
                    }
                    connection.Close();

                }
            }
            catch
            {
                throw;
            }
        }
        private static int Ile_Dni_Roboczych(List<Nieobecnosc> listaNieobecnosci)
        {
            int total = 0;
            foreach (var nieobecnosc in listaNieobecnosci)
            {
                DateTime absenceDate = new DateTime(nieobecnosc.rok, nieobecnosc.miesiac, nieobecnosc.dzien);
                if (absenceDate.DayOfWeek != DayOfWeek.Saturday && absenceDate.DayOfWeek != DayOfWeek.Sunday)
                {
                    total++;
                }
            }
            return total;
        }
        private static (TimeSpan, TimeSpan, TimeSpan, TimeSpan, TimeSpan, TimeSpan) Oblicz_Czas_Z_Dodatkiem(Dane_Dnia dane_Dni)
        {
            TimeSpan godzRozpPracy = dane_Dni.godz_rozp_pracy;
            TimeSpan godzZakonczPracy = dane_Dni.godz_zakoncz_pracy;
            double godzNadlPlatne50 = (double)dane_Dni.Godz_nadl_platne_z_dod_50;
            double godzNadlPlatne100 = (double)dane_Dni.Godz_nadl_platne_z_dod_100;

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

            return (new TimeSpan((int)startPodstawowy.TotalHours % 24, startPodstawowy.Minutes, startPodstawowy.Seconds),
                new TimeSpan((int)endPodstawowy.TotalHours % 24, endPodstawowy.Minutes, endPodstawowy.Seconds),
                new TimeSpan((int)startNadl50.TotalHours % 24, startNadl50.Minutes, startNadl50.Seconds),
                new TimeSpan((int)endNadl50.TotalHours % 24, endNadl50.Minutes, endNadl50.Seconds),
                new TimeSpan((int)startNadl100.TotalHours % 24, startNadl100.Minutes, startNadl100.Seconds),
                new TimeSpan((int)endNadl100.TotalHours % 24, endNadl100.Minutes, endNadl100.Seconds));
        }
        private static int Dodaj_Godz_Odbior_Do_Optimy(Karta_Pracy karta, SqlTransaction tran, SqlConnection connection)
        {
            int ilosc_wpisow = 0;
            DateTime baseDate = new DateTime(1899, 12, 30);
            var sqlInsertOdbNadg = @"
DECLARE @PRA_PraId INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId);
DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WHERE PPR_PraId = @PRA_PraId and PPR_Data = @DataInsert)
IF @EXISTSDZIEN is null
BEGIN
    BEGIN TRY
        INSERT INTO [CDN].[PracPracaDni]
                    ([PPR_PraId]
                    ,[PPR_Data]
                    ,[PPR_TS_Zal]
                    ,[PPR_TS_Mod]
                    ,[PPR_OpeModKod]
                    ,[PPR_OpeModNazwisko]
                    ,[PPR_OpeZalKod]
                    ,[PPR_OpeZalNazwisko]
                    ,[PPR_Zrodlo])
                VALUES
                    (@PRI_PraId
                    ,@DataInsert
                    ,GETDATE()
                    ,GETDATE()
                    ,'ADMIN'
                    ,'Administrator'
                    ,'ADMIN'
                    ,'Administrator'
                    ,0)
    END TRY
    BEGIN CATCH
    END CATCH
END

INSERT INTO CDN.PracPracaDniGodz
		(PGR_PprId,
		PGR_Lp,
		PGR_OdGodziny,
		PGR_DoGodziny,
		PGR_Strefa,
		PGR_DzlId,
		PGR_PrjId,
		PGR_Uwagi,
		PGR_OdbNadg)
	VALUES
		((select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId),
		1,
		DATEADD(MINUTE, 0, @GodzOdDate),
		DATEADD(MINUTE, 0, @GodzDoDate),
		@TypPracy,
		1,
		1,
		'',
		@TypNadg);";
            foreach (var dane_Dni in karta.dane_dni)
            {
                var Ilosc_Godzin = dane_Dni.Godz_Odbior;
                TimeSpan startGodz = TimeSpan.FromHours(8);
                TimeSpan endGodz = TimeSpan.FromHours(8) + TimeSpan.FromHours((double)dane_Dni.Godz_Odbior);
                DateTime godzOdDate = baseDate + startGodz;
                DateTime godzDoDate = baseDate + endGodz;
                bool duplicate = false;
                    using (SqlCommand cmd = new SqlCommand(@"
DECLARE @EXISTSDZIEN INT;
DECLARE @EXISTSDATA INT;
SET @EXISTSDZIEN = (SELECT COUNT(PPR_Data) FROM cdn.PracPracaDni WHERE PPR_PraId = @PRI_PraId AND PPR_Data = @DataInsert);
SET @EXISTSDATA = (
    SELECT COUNT(*)
    FROM CDN.PracPracaDniGodz 
    WHERE PGR_OdbNadg = 4
        AND PGR_Strefa = 2
        AND PGR_OdGodziny = DATEADD(MINUTE, 0, @GodzOdDate)
        AND PGR_DoGodziny = DATEADD(MINUTE, 0, @GodzDoDate)
        AND PGR_PprId = (SELECT PPR_PprId FROM cdn.PracPracaDni WHERE CAST(PPR_Data AS datetime) = @DataInsert AND PPR_PraId = @PRI_PraId)
);
SELECT CASE 
    WHEN @EXISTSDZIEN > 0 AND @EXISTSDATA > 0 THEN 1
    ELSE 0
END;", connection, tran))
                    {
                        cmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(karta.pracownik));
                        cmd.Parameters.AddWithValue("@TypPracy", 2);
                        cmd.Parameters.AddWithValue("@TypNadg", 4);
                        cmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                        cmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                        cmd.Parameters.AddWithValue("@DataInsert", DateTime.Parse($"{karta.rok}-{karta.miesiac:D2}-{dane_Dni.dzien:D2}"));
                        if ((int)cmd.ExecuteScalar() == 1)
                        {
                            duplicate = true;
                        }
                    }
                if (!duplicate)
                {
                    try
                    {
                        if (dane_Dni.Godz_Odbior > 0)
                        {
                            ilosc_wpisow++;

                            using (SqlCommand insertCmd = new SqlCommand(sqlInsertOdbNadg, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.Parse($"{karta.rok}-{karta.miesiac:D2}-{dane_Dni.dzien:D2}"));
                                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                insertCmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(karta.pracownik));
                                insertCmd.Parameters.AddWithValue("@TypPracy", 2); // podstawowy
                                insertCmd.Parameters.AddWithValue("@TypNadg", 4); // W.PŁ
                                if (Program.error_logger.Last_Mod_Osoba.Length > 20)
                                {
                                    insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 20));
                                }
                                else
                                {
                                    insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba);
                                }
                                if (Program.error_logger.Last_Mod_Osoba.Length > 50)
                                {
                                    insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 50));
                                }
                                else
                                {
                                    insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba);
                                }
                                insertCmd.Parameters.AddWithValue("@DataMod", Program.error_logger.Last_Mod_Time);
                                insertCmd.ExecuteScalar();
                            }
                        }
                    }
                    catch (SqlException ex)
                    {
                        tran.Rollback();
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    }
                    catch (FormatException)
                    {
                        continue;
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    }
                }
            }
            return ilosc_wpisow;
        }
        private static int Get_ID_Pracownika(Pracownik pracownik)
        {
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                try
                {
                    connection.Open();
                    using (SqlCommand getCmd = new SqlCommand(Program.sqlQueryGetPRI_PraId, connection))
                    {
                        getCmd.Parameters.AddWithValue("@Akronim ", pracownik.Akronim);
                        getCmd.Parameters.AddWithValue("@PracownikImieInsert", pracownik.Imie);
                        getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", pracownik.Nazwisko);
                        var result = getCmd.ExecuteScalar();
                        if (result != null)
                        {
                            return Convert.ToInt32(result);
                        }
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
            }
            return 0;
        }
    }
}
