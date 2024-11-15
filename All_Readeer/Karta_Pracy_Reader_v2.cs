using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace All_Readeer
{
    internal class Karta_Pracy_Reader_v2
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
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
                    else if (nazwa.ToLower() == "luty")
                    {
                        miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "marzec")
                    {
                        miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "kwiecień")
                    {
                        miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "maj")
                    {
                        miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "czerwiec")
                    {
                        miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "lipiec")
                    {
                        miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "sierpień")
                    {
                        miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "wrzesień")
                    {
                        miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "październik")
                    {
                        miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "listopad")
                    {
                        miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "grudzień")
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
        }
        private class CurrentPosition
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
        private string Last_Mod_Osoba = "";
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Optima_Connection_String = "";
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            Optima_Connection_String = NewConnectionString;
        }
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time)
        {
            try
            {
                Last_Mod_Osoba = last_Mod_Osoba;
                Last_Mod_Time = last_Mod_Time;
                List<Karta_Pracy> karty_pracy = [];
                CurrentPosition pozycja = new();
                Find_Karta(ref pozycja, worksheet);
                Karta_Pracy karta_pracy = new();
                karta_pracy.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                karta_pracy.nr_zakladki = Program.error_logger.Nr_Zakladki;
                Nieobecnosc nieobecnosc = new();
                nieobecnosc.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                nieobecnosc.nr_zakladki = Program.error_logger.Nr_Zakladki;
                Get_Header_Karta_Info(pozycja, worksheet, ref karta_pracy);

                Get_Dane_Dni(pozycja, worksheet, ref karta_pracy);
                karty_pracy.Add(karta_pracy);
                if(karty_pracy.Count > 0)
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
            }catch(Exception ex){
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        private void Find_Karta(ref CurrentPosition pozycja, IXLWorksheet worksheet)
        {
            pozycja.col = 2;
            bool found = false;
            try
            {
                foreach (var cell in worksheet.Column(pozycja.col).CellsUsed())
                {
                    if (cell.GetValue<string>().Equals("Dzień", StringComparison.OrdinalIgnoreCase))
                    {
                        pozycja.row = cell.Address.RowNumber;
                        found = true;
                        return;
                    }
                }
                if (!found)
                {
                    throw new Exception("Nie znaleziono słowa 'Dzień' w kolumnie.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw new Exception("Nie znaleziono słowa 'Dzień' w kolumnie.");
            }
        }
        private void Get_Header_Karta_Info(CurrentPosition StartKarty, IXLWorksheet worksheet, ref Karta_Pracy karta_pracy)
        {
            //wczytaj date
            var dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + 12).GetValue<string>().Trim();
            if (dane.EndsWith("r"))
            {
                dane = dane.Substring(0, dane.Length - 1).Trim();
            }
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "data", StartKarty.col + 12, StartKarty.row - 3, "Brak daty w pliku");
                Console.WriteLine(Program.error_logger.Get_Error_String());
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            if (karta_pracy.Set_Data(dane) == 1) {
                if (dane.Split(" ").Length == 2)
                {
                    var ndata = dane.Split(" ");
                    if (ndata[0].ToLower() == "pażdziernik")
                    {
                        ndata[0] = "październik";
                    }
                    try
                    {
                        karta_pracy.Set_Miesiac(ndata[0]);
                        karta_pracy.rok = int.Parse(Regex.Replace(ndata[1], @"\D", ""));
                    }
                    catch
                    {
                        karta_pracy.Set_Miesiac("Zle dane");
                        karta_pracy.rok = int.Parse(Regex.Replace(ndata[1], @"\D", ""));
                    }
                    if (karta_pracy.miesiac == 0)
                    {
                        Program.error_logger.New_Error(dane, "data", StartKarty.col + 12, StartKarty.row - 3, "Zły format daty w pliku");
                        Console.WriteLine(Program.error_logger.Get_Error_String());
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
                else
                {
                    Program.error_logger.New_Error(dane, "data", StartKarty.col + 12, StartKarty.row - 3, "Zły format daty w pliku");
                    Console.WriteLine(Program.error_logger.Get_Error_String());
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }

            //wczytaj nazwisko i imie
            dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col).GetValue<string>().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Nie wykryto nazwiska i imienia w pliku");
                Console.WriteLine(Program.error_logger.Get_Error_String());
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            if(dane.Contains("KARTA PRACY:"))
            {
                dane = dane.Replace("KARTA PRACY:", "").Trim();
            }
            if (string.IsNullOrEmpty(dane))
            {
                dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col + 1).GetValue<string>().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(dane))
                {
                    dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col + 2).GetValue<string>().Trim().Replace("  ", " ");
                    if (string.IsNullOrEmpty(dane))
                    {
                        Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col - 2, StartKarty.row, "Zły format pola nazwisko i imie");
                        Console.WriteLine(Program.error_logger.Get_Error_String());
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                    else
                    {
                        karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                        karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[1];
                    }
                }
                else
                {
                    karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                    karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[1];
                }
            }
            else
            {
                karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[1];
            }
            if (karta_pracy.pracownik.Nazwisko == null || karta_pracy.pracownik.Imie == null)
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col - 2, StartKarty.row, "Zły format pola nazwisko i imie");
                Console.WriteLine(Program.error_logger.Get_Error_String());
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private void Get_Dane_Dni(CurrentPosition StartKarty, IXLWorksheet worksheet, ref Karta_Pracy karta_pracy)
        {
            StartKarty.row += 3;
            var NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
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
                    Console.WriteLine(Program.error_logger.Get_Error_String());
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                //try get nieobecność:
                try
                {
                    var danei = worksheet.Cell(StartKarty.row, StartKarty.col + 3).GetValue<string>();
                    if (!string.IsNullOrEmpty(danei.Trim()))
                    {
                        Nieobecnosc nieobecnosc = new();
                        if (RodzajNieobecnosci.TryParse(danei.ToUpper(), out RodzajNieobecnosci Rnieobecnosc))
                        {
                            nieobecnosc.rodzaj_absencji = Rnieobecnosc;
                            nieobecnosc.pracownik = karta_pracy.pracownik;
                            nieobecnosc.rok = karta_pracy.rok;
                            nieobecnosc.miesiac = karta_pracy.miesiac;
                            nieobecnosc.dzien = dzien.dzien;
                        }
                        else
                        {
                            Console.WriteLine($"Nieprawidłowy kod nieobecności: {danei} w pliku {Program.error_logger.Nazwa_Pliku} w zakladce {Program.error_logger.Nr_Zakladki}");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        karta_pracy.ListaNieobecnosci.Add(nieobecnosc);
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw new Exception(ex.Message);
                }
                // godz rozpoczecia
                try
                {
                    var danei = worksheet.Cell(StartKarty.row, StartKarty.col + 1).GetValue<string>().Trim().Split(' ')[1];
                    danei = danei.Replace('.', ':');
                    if (string.IsNullOrEmpty(danei))
                    {
                        //ErrorLogger_v2.New_Error(danei, "godziny pozpoczęcia pracy", StartKarty.col + 1, StartKarty.row, "Brak wpisanej godziny pozpoczęcia pracy");
                        //throw new Exception(ErrorLogger_v2.Get_Error_String());
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
                        continue;
                    }
                    if (TimeSpan.TryParse(danei, out TimeSpan czasRozpoczecia))
                    {
                        dzien.godz_rozp_pracy = czasRozpoczecia;
                    }
                    else
                    {
                        //here try to solve times like 07:59:60 xdd
                        if (danei.Split(':').Count() == 3)
                        {
                            var parts = danei.Split(':');

                            if (int.TryParse(parts[0], out int hours) &&
                                int.TryParse(parts[1], out int minutes) &&
                                int.TryParse(parts[2], out int seconds))
                            {
                                if (seconds >= 60)
                                {
                                    seconds -= 60;
                                    minutes += 1;
                                }
                                if (minutes >= 60)
                                {
                                    minutes -= 60;
                                    hours += 1;
                                }
                                hours %= 24;
                                dzien.godz_rozp_pracy = new TimeSpan(hours, minutes, seconds);
                            }

                        }
                        else
                        {
                            Program.error_logger.New_Error(danei, "godz_rozp_pracy", StartKarty.col + 1, StartKarty.row);
                            Console.WriteLine(Program.error_logger.Get_Error_String() + " (Zły foramt czasu)");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                }
                catch
                {
                    var danei = worksheet.Cell(StartKarty.row, StartKarty.col + 1).GetValue<string>().Trim();
                    danei = danei.Replace('.', ':');
                    if (string.IsNullOrEmpty(danei))
                    {
                        //ErrorLogger_v2.New_Error(danei, "godziny pozpoczęcia pracy", StartKarty.col + 1, StartKarty.row, "Brak wpisanej godziny pozpoczęcia pracy");
                        //throw new Exception(ErrorLogger_v2.Get_Error_String());
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
                        continue;
                    }
                    if (TimeSpan.TryParse(danei, out TimeSpan czasRozpoczecia))
                    {
                        dzien.godz_rozp_pracy = czasRozpoczecia;
                    }
                    else
                    {
                        //here try to solve times like 07:59:60 xdd
                        if (danei.Split(':').Count() == 3)
                        {
                            var parts = danei.Split(':');

                            if (int.TryParse(parts[0], out int hours) &&
                                int.TryParse(parts[1], out int minutes) &&
                                int.TryParse(parts[2], out int seconds))
                            {
                                if (seconds >= 60)
                                {
                                    seconds -= 60;
                                    minutes += 1;
                                }
                                if (minutes >= 60)
                                {
                                    minutes -= 60;
                                    hours += 1;
                                }
                                hours %= 24;
                                dzien.godz_rozp_pracy = new TimeSpan(hours, minutes, seconds);
                            }

                        }
                        else
                        {
                            Program.error_logger.New_Error(danei, "godz_rozp_pracy", StartKarty.col + 1, StartKarty.row);
                            Console.WriteLine(Program.error_logger.Get_Error_String() + " (Zły foramt czasu)");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                }
                // godz zakonczenia
                try
                {
                    var danei = worksheet.Cell(StartKarty.row, StartKarty.col + 2).GetValue<string>().Trim().Split(' ')[1];
                    danei = danei.Replace('.', ':');
                    if (string.IsNullOrEmpty(danei))
                    {
                        //ErrorLogger_v2.New_Error(danei, "godziny zakończenia pracy", StartKarty.col + 2, StartKarty.row, "Brak wpisanej godziny zakonczenia pracy");
                        //throw new Exception(ErrorLogger_v2.Get_Error_String());
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
                        continue;
                    }
                    if (TimeSpan.TryParse(danei, out TimeSpan czasZakonczenia))
                    {
                        dzien.godz_zakoncz_pracy = czasZakonczenia;
                    }
                    else
                    {
                        //here try to solve times like 07:59:60 xdd
                        if (danei.Split(':').Count() == 3)
                        {
                            var parts = danei.Split(':');

                            if (int.TryParse(parts[0], out int hours) &&
                                int.TryParse(parts[1], out int minutes) &&
                                int.TryParse(parts[2], out int seconds))
                            {
                                if (seconds >= 60)
                                {
                                    seconds -= 60;
                                    minutes += 1;
                                }
                                if (minutes >= 60)
                                {
                                    minutes -= 60;
                                    hours += 1;
                                }
                                hours %= 24;
                                dzien.godz_zakoncz_pracy = new TimeSpan(hours, minutes, seconds);
                            }

                        }
                        else
                        {
                            Program.error_logger.New_Error(danei, "godz_zakoncz_pracy", StartKarty.col +2, StartKarty.row);
                            Console.WriteLine(Program.error_logger.Get_Error_String() + " (Zły foramt czasu)");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                }
                catch
                {
                    var danei = worksheet.Cell(StartKarty.row, StartKarty.col + 2).GetValue<string>().Trim();
                    danei = danei.Replace('.', ':');
                    if (string.IsNullOrEmpty(danei))
                    {
                        //ErrorLogger_v2.New_Error(danei, "godziny zakończenia pracy", StartKarty.col + 2, StartKarty.row, "Brak wpisanej godziny zakonczenia pracy");
                        //throw new Exception(ErrorLogger_v2.Get_Error_String());
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
                        continue;
                    }
                    if (TimeSpan.TryParse(danei, out TimeSpan czasZakonczenia))
                    {
                        dzien.godz_zakoncz_pracy = czasZakonczenia;
                        if (dzien.godz_zakoncz_pracy == TimeSpan.Zero)
                        {
                            dzien.godz_zakoncz_pracy = TimeSpan.FromHours(24);
                        }
                    }
                    else
                    {
                        //here try to solve times like 07:59:60 xdd
                        if (danei.Split(':').Count() == 3)
                        {
                            var parts = danei.Split(':');

                            if (int.TryParse(parts[0], out int hours) &&
                                int.TryParse(parts[1], out int minutes) &&
                                int.TryParse(parts[2], out int seconds))
                            {
                                if (seconds >= 60)
                                {
                                    seconds -= 60;
                                    minutes += 1;
                                }
                                if (minutes >= 60)
                                {
                                    minutes -= 60;
                                    hours += 1;
                                }
                                hours %= 24;
                                dzien.godz_zakoncz_pracy = new TimeSpan(hours, minutes, seconds);
                            }

                        }
                        else
                        {
                            Program.error_logger.New_Error(danei, "godz_zakoncz_pracy", StartKarty.col + 2, StartKarty.row);
                            Console.WriteLine(Program.error_logger.Get_Error_String() + " (Zły foramt czasu)");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                    }
                }
                if(dzien.godz_rozp_pracy != TimeSpan.Zero && dzien.godz_zakoncz_pracy != TimeSpan.Zero)
                {
                    karta_pracy.dane_dni.Add(dzien);
                }
                StartKarty.row++;
                NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
            }
        }
        private void Wpierdol_Obecnosci_do_Optimy(Karta_Pracy karta, SqlTransaction tran, SqlConnection connection)
        {
            foreach (var dane_Dni in karta.dane_dni)
            {
                try
                {
                    // jak praca po północy to na next dzien przeniesc
                    if (dane_Dni.godz_zakoncz_pracy < dane_Dni.godz_rozp_pracy)
                    {
                        // insert godziny przed północą
                        using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{karta.rok}-{karta.miesiac:D2}-{dane_Dni.dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                            DateTime baseDate = new DateTime(1899, 12, 30);
                            DateTime godzOdDate = baseDate + dane_Dni.godz_rozp_pracy;
                            DateTime godzDoDate = baseDate + TimeSpan.FromHours(24);
                            insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                            insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                            double czasPrzepracowanyInsert = (TimeSpan.FromHours(24) - dane_Dni.godz_rozp_pracy).TotalHours;
                            insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                            insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                            insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                            insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                            insertCmd.ExecuteScalar();
                        }
                        // insert po północy
                        using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                        {
                            var data = new DateTime(karta.rok, karta.miesiac, dane_Dni.dzien).AddDays(1);
                            insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{data:yyyy-MM-dd}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                            DateTime baseDate = new DateTime(1899, 12, 30);
                            DateTime godzOdDate = baseDate + TimeSpan.FromHours(0);
                            DateTime godzDoDate = baseDate + dane_Dni.godz_zakoncz_pracy;
                            insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                            insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                            double czasPrzepracowanyInsert = dane_Dni.godz_zakoncz_pracy.TotalHours;
                            insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                            insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                            insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                            insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                            insertCmd.ExecuteScalar();
                        }
                    }
                    else
                    {
                        using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{karta.rok}-{karta.miesiac:D2}-{dane_Dni.dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));

                            DateTime dataBazowa = new DateTime(1899, 12, 30);
                            DateTime godzOdDate = dataBazowa + dane_Dni.godz_rozp_pracy;
                            DateTime godzDoDate = dataBazowa + dane_Dni.godz_zakoncz_pracy;
                            insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                            insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                            insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", (dane_Dni.godz_zakoncz_pracy - dane_Dni.godz_rozp_pracy).TotalHours);
                            insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", (dane_Dni.godz_zakoncz_pracy - dane_Dni.godz_rozp_pracy).TotalHours);
                            insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                            insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                            insertCmd.ExecuteScalar();
                        }
                    }
                }
                catch (SqlException ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
            }
        }
        private void Wjeb_Nieobecnosci_do_Optimy(List<Nieobecnosc> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
        {
            List<List<Nieobecnosc>> Nieobecnosci = Podziel_Niobecnosci_Na_Osobne(ListaNieobecności);
            foreach (var ListaNieo in Nieobecnosci)
                {
                var dni_robocze = Ile_Dni_Roboczych(ListaNieo);
                var dni_calosc = ListaNieo.Count;

                try
                {
                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertNieObecnoŚciDoOptimy, connection, tran))
                    {
                        DateTime dataBazowa = new DateTime(1899, 12, 30);
                        var nazwa_nieobecnosci = Dopasuj_Nieobecnosc(ListaNieo[0].rodzaj_absencji);
                        if (string.IsNullOrEmpty(nazwa_nieobecnosci))
                        {
                            Program.error_logger.New_Custom_Error($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                            var e = new Exception($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                            e.Data["Kod"] = 42069;
                            throw e;
                        }
                        DateTime dataniobecnoscistart = new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien);
                        DateTime dataniobecnosciend = new DateTime(ListaNieo[ListaNieo.Count-1].rok, ListaNieo[ListaNieo.Count-1].miesiac, ListaNieo[ListaNieo.Count-1].dzien);
                        int przyczyna = Dopasuj_Przyczyne(ListaNieo[0].rodzaj_absencji);
                        insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", ListaNieo[0].pracownik.Nazwisko);
                        insertCmd.Parameters.AddWithValue("@PracownikImieInsert", ListaNieo[0].pracownik.Imie);
                        insertCmd.Parameters.AddWithValue("@NazwaNieobecnosci", nazwa_nieobecnosci);
                        insertCmd.Parameters.AddWithValue("@DniPracy", dni_robocze);
                        insertCmd.Parameters.AddWithValue("@DniKalendarzowe", dni_calosc);
                        insertCmd.Parameters.AddWithValue("@Przyczyna", przyczyna);
                        insertCmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = dataniobecnoscistart;
                        insertCmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = dataBazowa;
                        insertCmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = dataniobecnosciend;
                        if (Last_Mod_Osoba.Length > 20)
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Last_Mod_Osoba.Substring(0, 20));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Last_Mod_Osoba);
                        }
                        if (Last_Mod_Osoba.Length > 50)
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Last_Mod_Osoba.Substring(0, 50));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Last_Mod_Osoba);
                        }
                        insertCmd.Parameters.AddWithValue("@DataMod", Last_Mod_Time);
                        insertCmd.ExecuteScalar();
                    }
                }
                catch (SqlException ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    if (ex.Data.Contains("kod") && ex.Data["kod"] is int kod && kod == 42069)
                    {
                        throw;
                    }
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
            }

        }
        private string Dopasuj_Nieobecnosc(RodzajNieobecnosci rodzaj)
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
        private int Dopasuj_Przyczyne(RodzajNieobecnosci rodzaj)
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
        private List<List<Nieobecnosc>> Podziel_Niobecnosci_Na_Osobne(List<Nieobecnosc> listaNieobecnosci)
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
        private void Dodaj_Dane_Do_Optimy(Karta_Pracy karta)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Optima_Connection_String))
                {
                    connection.Open();
                    SqlTransaction tran = connection.BeginTransaction();
                    Wpierdol_Obecnosci_do_Optimy(karta, tran, connection);
                    Wjeb_Nieobecnosci_do_Optimy(karta.ListaNieobecnosci, tran, connection);
                    tran.Commit();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno nieobecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.WriteLine($"Poprawnie dodawno obecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.ForegroundColor = ConsoleColor.White;
                    connection.Close();
                }
            }
            catch
            {
                throw;
            }
        }
        private int Ile_Dni_Roboczych(List<Nieobecnosc> listaNieobecnosci)
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
    }
}
