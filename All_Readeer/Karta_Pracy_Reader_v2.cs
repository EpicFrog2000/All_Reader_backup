using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

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
        private string File_Path = "";
        private string Last_Mod_Osoba = "";
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Connection_String = "";
        private string Optima_Connection_String = "";
        private (string, DateTime) Get_File_Meta_Info()
        {
            try
            {
                using (var workbook = new XLWorkbook(File_Path))
                {
                    string lastModifiedBy = workbook.Properties.LastModifiedBy!;
                    DateTime lastWriteTime = File.GetLastWriteTime(File_Path);
                    return (lastModifiedBy, lastWriteTime);
                }
            }
            catch (Exception ex)
            {
                Program.error_logger.New_Custom_Error(ex.Message);
                Console.WriteLine($"Error: {ex.Message}");
                return ("Error", DateTime.Now);
            }
        }
        public void Set_File_Path(string New_File_Path)
        {
            if (string.IsNullOrEmpty(New_File_Path))
            {
                Console.WriteLine("error: Empty File Path");
                return;
            }
            File_Path = New_File_Path;
        }
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            Optima_Connection_String = NewConnectionString;
        }
        public void Process()
        {
            List<Karta_Pracy> karty_pracy = ReadXlsx();
            // wpierdol dane do bazy danych
            List<Pracownik> pracownicy = karty_pracy.Select(k => k.pracownik).Distinct().ToList();
            try
            {
                //Insert_Pracownicy_To_Db(pracownicy);
                foreach (var karta in karty_pracy)
                {
                    try
                    {
                        //int id = Insert_Karta_To_Db(karta);
                        //Insert_Dni_To_Db(id, karta.dane_dni);
                        Dodaj_Dane_Do_Optimy(karta);
                    }
                    catch (Exception ex) {
                        Console.WriteLine(ex.Message);
                    }
                }
            } catch (Exception ex) {
                Console.WriteLine(ex.Message);
            }
        }
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time)
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
            foreach (var karta in karty_pracy)
            {
                try
                {
                    Dodaj_Dane_Do_Optimy(karta);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw new Exception(ex.Message);
                }
            }
        }
        private List<Karta_Pracy> ReadXlsx()
        {
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            Program.error_logger.Nazwa_Pliku = File_Path;
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            List<Karta_Pracy> karty_pracy = [];
            karty_pracy.Clear();
            using (var workbook = new XLWorkbook(File_Path))
            {
                CurrentPosition pozycja = new();
                for (int i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    try
                    {
                        Program.error_logger.Nr_Zakladki = i;
                        var worksheet = workbook.Worksheet(i);
                        Find_Karta(ref pozycja, worksheet);
                        Karta_Pracy karta_pracy = new();
                        karta_pracy.nazwa_pliku = File_Path;
                        karta_pracy.nr_zakladki = i;
                        Nieobecnosc nieobecnosc = new();
                        nieobecnosc.nazwa_pliku = File_Path;
                        nieobecnosc.nr_zakladki = i;
                        Get_Header_Karta_Info(pozycja, worksheet, ref karta_pracy);
                        Get_Dane_Dni(pozycja, worksheet, ref karta_pracy);
                        karty_pracy.Add(karta_pracy);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return karty_pracy;
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
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "data", StartKarty.col + 12, StartKarty.row - 3, "Brak daty w pliku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            if (karta_pracy.Set_Data(dane) == 1) {
                Program.error_logger.New_Error(dane, "data", StartKarty.col + 12, StartKarty.row - 3, "Zły format daty w pliku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            //wczytaj nazwisko i imie
            dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col).GetValue<string>().Trim().Replace("  ", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Nie wykryto nazwiska i imienia w pliku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            karta_pracy.pracownik.Nazwisko = dane.Split(':')[1].Trim().Split(' ')[0];
            karta_pracy.pracownik.Imie = dane.Split(':')[1].Trim().Split(' ')[1];
            if (karta_pracy.pracownik.Nazwisko == null || karta_pracy.pracownik.Imie == null)
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Zły format pola nazwisko i imie");
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
                }
                else
                {
                    Program.error_logger.New_Error(NrDnia, "dzien", StartKarty.col, StartKarty.row, "Błędny nr dnia");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                //try get nieobecność:
                try
                {
                    var danei = worksheet.Cell(StartKarty.row, StartKarty.col + 3).GetValue<string>();
                    if (!string.IsNullOrEmpty(danei))
                    {
                        Nieobecnosc nieobecnosc = new();
                        if (Enum.TryParse(danei, out RodzajNieobecnosci Rnieobecnosc))
                        {
                            nieobecnosc.rodzaj_absencji = Rnieobecnosc;
                            nieobecnosc.pracownik = karta_pracy.pracownik;
                            nieobecnosc.rok = karta_pracy.rok;
                            nieobecnosc.miesiac = karta_pracy.miesiac;
                            nieobecnosc.dzien = dzien.dzien;
                        }
                        else
                        {
                            Program.error_logger.New_Error(danei, "kod nieobecnosci", StartKarty.col + 3, StartKarty.row, "Nieprawidłowy kod nieobecności");
                            Console.WriteLine($"Nieprawidłowy kod nieobecności: {danei}");
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
                        Program.error_logger.New_Error(danei, "godziny pozpoczęcia pracy", StartKarty.col + 1, StartKarty.row, "Nieprawidłowy format czasu pracy");
                        throw new Exception(Program.error_logger.Get_Error_String());
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
                        Program.error_logger.New_Error(danei, "godziny pozpoczęcia pracy", StartKarty.col + 1, StartKarty.row, "Nieprawidłowy format czasu pracy");
                        throw new Exception(Program.error_logger.Get_Error_String());
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
                        Program.error_logger.New_Error(danei, "godziny zakończenia pracy", StartKarty.col + 2, StartKarty.row, "Nieprawidłowy format czasu pracy.");
                        throw new Exception(Program.error_logger.Get_Error_String());
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
                        Program.error_logger.New_Error(danei, "godziny zakończenia pracy", StartKarty.col + 2, StartKarty.row, "Nieprawidłowy format czasu pracy.");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }

                // Liczba godz. przepracowanych
                var dane = worksheet.Cell(StartKarty.row, StartKarty.col + 7).GetValue<string>().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    Program.error_logger.New_Error(dane, "liczba godz przepracowanych", StartKarty.col + 7, StartKarty.row, "Brak wpisanej liczba godz przepracowanych");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                if (decimal.TryParse(dane, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"), out decimal liczba))
                {
                    dzien.liczba_godz_przepracowanych = Math.Abs(liczba);
                }
                else
                {
                    Program.error_logger.New_Error(dane, "liczba godz przepracowanych", StartKarty.col + 7, StartKarty.row, "Zly format liczba godz przepracowanych");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // Praca wg grafiku
                dane = worksheet.Cell(StartKarty.row, StartKarty.col + 8).GetValue<string>().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    //ErrorLogger_v2.New_Error(dane, "praca wg grafiku", StartKarty.col + 8, StartKarty.row, "Brak wpisanej praca wg grafiku");
                    //throw new Exception(ErrorLogger_v2.Get_Error_String());
                } else if (decimal.TryParse(dane, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"), out liczba))
                {
                    dzien.praca_wg_grafiku = Math.Abs(liczba);
                }
                else
                {
                    Program.error_logger.New_Error(dane, "praca wg grafiku", StartKarty.col + 8, StartKarty.row, "Zly format praca wg grafiku");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                //Godziny nadl. płatne z dod. 50%
                dane = worksheet.Cell(StartKarty.row, StartKarty.col + 9).GetValue<string>().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    //ErrorLogger_v2.New_Error(dane, "Godz nadl platne z dod 50", StartKarty.col + 9, StartKarty.row, "Brak wpisanej Godz nadl platne z dod 50");
                    //throw new Exception(ErrorLogger_v2.Get_Error_String());
                } else if (decimal.TryParse(dane, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"), out liczba))
                {
                    dzien.Godz_nadl_platne_z_dod_50 = Math.Abs(liczba);
                }
                else
                {
                    Program.error_logger.New_Error(dane, "Godz nadl platne z dod 50", StartKarty.col + 9, StartKarty.row, "Zly format Godz nadl platne z dod 50");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                //Godziny nadl. płatne z dod. 100%
                dane = worksheet.Cell(StartKarty.row, StartKarty.col + 10).GetValue<string>().Trim();
                if (string.IsNullOrEmpty(dane))
                {
                    //ErrorLogger_v2.New_Error(dane, "Godz nadl platne z dod 100", StartKarty.col + 10, StartKarty.row, "Brak wpisanej Godz nadl platne z dod 100");
                    //throw new Exception(ErrorLogger_v2.Get_Error_String());
                } else if (decimal.TryParse(dane, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.GetCultureInfo("pl-PL"), out liczba))
                {
                    dzien.Godz_nadl_platne_z_dod_100 = Math.Abs(liczba);
                }
                else
                {
                    Program.error_logger.New_Error(dane, "Godz nadl platne z dod 100", StartKarty.col + 10, StartKarty.row, "Zly format Godz nadl platne z dod 100");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                karta_pracy.dane_dni.Add(dzien);
                StartKarty.row++;
                NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
            }
        }
        public void Set_Db_Tables_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }
        private void Insert_Pracownicy_To_Db(List<Pracownik> pracownicy)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    foreach (var pracownik in pracownicy)
                    {
                        string checkQuery = "SELECT COUNT(1) FROM Karta_Pracy_Kucharska_Pracownicy WHERE Imie = @Imie AND Nazwisko = @Nazwisko";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection, tran))
                        {
                            checkCmd.Parameters.AddWithValue("@Imie", pracownik.Imie);
                            checkCmd.Parameters.AddWithValue("@Nazwisko", pracownik.Nazwisko);

                            int count = (int)checkCmd.ExecuteScalar();

                            if (count == 0)
                            {
                                string insertQuery = "INSERT INTO Karta_Pracy_Kucharska_Pracownicy (Imie, Nazwisko, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) VALUES (@Imie, @Nazwisko, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Os)";
                                using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                                {
                                    insertCmd.Parameters.AddWithValue("@Imie", pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@Nazwisko", pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);
                                    insertCmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    tran.Commit();
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
        }
        private int Insert_Karta_To_Db(Karta_Pracy karta)
        {
            int insertedId = -1;
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    var Id_Pracownika = 0;
                    string selectQuery = "SELECT Id_Pracownika FROM Karta_Pracy_Kucharska_Pracownicy WHERE Imie = @Imie AND Nazwisko = @Nazwisko;";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, connection, tran))
                    {
                        selectCmd.Parameters.Add("@Imie", SqlDbType.NVarChar).Value = karta.pracownik.Imie;
                        selectCmd.Parameters.Add("@Nazwisko", SqlDbType.NVarChar).Value = karta.pracownik.Nazwisko;
                        object result = selectCmd.ExecuteScalar();
                        Id_Pracownika = Convert.ToInt32(result);
                    }
                    string insertQuery = @"INSERT INTO Karty_Pracy_Kucharska (
                                    Id_Pracownika,
                                    Miesiac,
                                    Rok,
                                    Ostatnia_Modyfikacja_Data,
                                    Ostatnia_Modyfikacja_Osoba
                                )
                                VALUES(
                                    @Id_Pracownika,
                                    @Miesiac,
                                    @Rok,
                                    @Ostatnia_Modyfikacja_Data,
                                    @Ostatnia_Modyfikacja_Os
                                ); SELECT SCOPE_IDENTITY();";

                    using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                    {
                        insertCmd.Parameters.Add("@Id_Pracownika", SqlDbType.Int).Value = Id_Pracownika;
                        insertCmd.Parameters.Add("@Miesiac", SqlDbType.Int).Value = karta.miesiac;
                        insertCmd.Parameters.Add("@Rok", SqlDbType.Int).Value = karta.rok;
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);

                        insertedId = Convert.ToInt32(insertCmd.ExecuteScalar());
                        tran.Commit();
                    }
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
            return insertedId;
        }
        private void Insert_Dni_To_Db(int Id_Karty, List<Dane_Dnia> dni)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    string insertQuery = "INSERT INTO Karty_Pracy_Kucharska_Dane_Dni (Id_Karty, Dzien, Godzina_Rozpoczęcia_Pracy, Godzina_Zakonczenia_Pracy, Czas_Faktyczny_Przepracowany, Praca_WG_Grafiku, Ilosc_Godzin_Z_Dodatkiem_50, Ilosc_Godzin_Z_Dodatkiem_100, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                            "VALUES (@Id_Karty, @Dzien, @Godzina_Rozpoczecia, @Godzina_Zakonczenia_Pracy, @Czas_Faktyczny_Przepracowany, @Praca_WG_Grafiku, @Ilosc_Godzin_Z_Dodatkiem_50, @Ilosc_Godzin_Z_Dodatkiem_100, @Ostatnia_Modyfikacja_Data, @Ostatnia_Modyfikacja_Os);";

                    foreach (var dzień in dni)
                    {
                        using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@Id_Karty", Id_Karty);
                            insertCmd.Parameters.AddWithValue("@Dzien", dzień.dzien);
                            insertCmd.Parameters.AddWithValue("@Godzina_Rozpoczecia", dzień.godz_rozp_pracy);
                            insertCmd.Parameters.AddWithValue("@Godzina_Zakonczenia_Pracy", dzień.godz_zakoncz_pracy);
                            insertCmd.Parameters.AddWithValue("@Czas_Faktyczny_Przepracowany", dzień.liczba_godz_przepracowanych);
                            insertCmd.Parameters.AddWithValue("@Praca_WG_Grafiku", dzień.praca_wg_grafiku);
                            insertCmd.Parameters.AddWithValue("@Ilosc_Godzin_Z_Dodatkiem_50", dzień.Godz_nadl_platne_z_dod_50);
                            insertCmd.Parameters.AddWithValue("@Ilosc_Godzin_Z_Dodatkiem_100", dzień.Godz_nadl_platne_z_dod_100);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);
                            insertCmd.ExecuteNonQuery();
                        }
                    }
                    tran.Commit();
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
        }
        private void Wpierdol_Obecnosci_do_Optimy(Karta_Pracy karta, SqlTransaction tran, SqlConnection connection)
        {
            var sqlQuery = $@"
DECLARE @id int;

-- dodawaina pracownika do pracx i init pracpracdni
DECLARE @PRI_PraId INT = (SELECT DISTINCT PRI_PraId FROM CDN.Pracidx where PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Imie1 = @PracownikImieInsert and PRI_Typ = 1);

IF @PRI_PraId IS NULL
BEGIN
    DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie: ' + @PracownikNazwiskoInsert + ' ' + @PracownikImieInsert;
    THROW 50000, @ErrorMessage, 1;
END


DECLARE @EXISTSPRACTEST INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId)

IF @EXISTSPRACTEST IS NULL
BEGIN
INSERT INTO [CDN].[PracKod]
        ([PRA_Kod]
        ,[PRA_Archiwalny]
        ,[PRA_Nadrzedny]
        ,[PRA_EPEmail]
        ,[PRA_EPTelefon]
        ,[PRA_EPNrPokoju]
        ,[PRA_EPDostep]
        ,[PRA_HasloDoWydrukow])
    VALUES
        (@PRI_PraId
        ,0
        ,0
        ,''
        ,''
        ,''
        ,0
        ,'')
END

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

SET @id = (select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId);

-- DODANIE GODZIN W NORMIE
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
	(@id,
	1,
	DATEADD(MINUTE, 0, @GodzOdDate),
	DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate),
	2,
	1,
	1,
	'',
	1);

-- DODANIE NADGODZIN
IF(@CzasPrzepracowanyInsert > @PracaWgGrafikuInsert)
BEGIN

IF(@Godz_dod_50 > 0)
BEGIN
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
				(@id,
				1,
				DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate),
				DATEADD(MINUTE, 60 * @Godz_dod_50, DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate)),
				4,
				1,
				1,
				'',
				4);
	SET @CzasPrzepracowanyInsert = @CzasPrzepracowanyInsert - @Godz_dod_50;
END

IF(@CzasPrzepracowanyInsert > @PracaWgGrafikuInsert)
BEGIN
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
				(@id,
				1,
				DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate),
				@GodzDoDate,
				4,
				1,
				1,
				'',
				4);
END
END";
            foreach (var dane_Dni in karta.dane_dni)
            {
                try
                {
                    // jak praca po północy to na next dzien przeniesc
                    if (dane_Dni.godz_zakoncz_pracy < dane_Dni.godz_rozp_pracy)
                    {
                        // insert godziny przed północą
                        using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                        using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                        using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    tran.Rollback();
                    var e = new Exception(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                    throw e;
                }
            }
        }
        private void Wjeb_Nieobecnosci_do_Optimy(List<Nieobecnosc> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
        {
            var sqlQuery = @$"
DECLARE @PRACID INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRACID IS NULL
BEGIN
    DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie: ' + @PracownikNazwiskoInsert + ' ' + @PracownikImieInsert;
    THROW 50000, @ErrorMessage, 1;
END

DECLARE @TNBID INT = (select TNB_TnbId from cdn.TypNieobec WHERE TNB_Nazwa = @NazwaNieobecnosci)

INSERT INTO [CDN].[PracNieobec]
           ([PNB_PraId]
           ,[PNB_TnbId]
           ,[PNB_TyuId]
           ,[PNB_NaPodstPoprzNB]
           ,[PNB_OkresOd]
           ,[PNB_Seria]
           ,[PNB_Numer]
           ,[PNB_OkresDo]
           ,[PNB_Opis]
           ,[PNB_Rozliczona]
           ,[PNB_RozliczData]
           ,[PNB_ZwolFPFGSP]
           ,[PNB_UrlopNaZadanie]
           ,[PNB_Przyczyna]
           ,[PNB_DniPracy]
           ,[PNB_DniKalend]
           ,[PNB_Calodzienna]
           ,[PNB_ZlecZasilekPIT]
           ,[PNB_PracaRodzic]
           ,[PNB_Dziecko]
           ,[PNB_OpeZalId]
           ,[PNB_StaZalId]
           ,[PNB_TS_Zal]
           ,[PNB_TS_Mod]
           ,[PNB_OpeModKod]
           ,[PNB_OpeModNazwisko]
           ,[PNB_OpeZalKod]
           ,[PNB_OpeZalNazwisko]
           ,[PNB_Zrodlo])
     VALUES
           (@PRACID
           ,@TNBID
           ,99999
           ,0
           ,@DataOd
           ,''
           ,''
           ,@DataDo
           ,''
           ,0
           ,@BaseDate
           ,0
           ,0
           ,@Przyczyna
           ,@DniPracy
           ,@DniKalendarzowe
           ,1
           ,0
           ,0
           ,''
           ,1
           ,1
           ,@DataMod
           ,@DataMod
           ,@ImieMod
           ,@NazwiskoMod
           ,@ImieMod
           ,@NazwiskoMod
           ,0);";
            List<List<Nieobecnosc>> Nieobecnosci = Podziel_Niobecnosci_Na_Osobne(ListaNieobecności);
            foreach (var ListaNieo in Nieobecnosci)
                {
                var dni_robocze = Ile_Dni_Roboczych(ListaNieo);
                var dni_calosc = ListaNieo.Count;

                try
                {
                    using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
                    {
                        DateTime dataBazowa = new DateTime(1899, 12, 30);
                        var nazwa_nieobecnosci = Dopasuj_Nieobecnosc(ListaNieo[0].rodzaj_absencji);
                        if (string.IsNullOrEmpty(nazwa_nieobecnosci))
                        {
                            Program.error_logger.New_Custom_Error($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                            var e = new Exception($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                            e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
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
                        insertCmd.Parameters.AddWithValue("@ImieMod", Last_Mod_Osoba);
                        insertCmd.Parameters.AddWithValue("@NazwiskoMod", Last_Mod_Osoba);
                        insertCmd.Parameters.AddWithValue("@DataMod", Last_Mod_Time);
                        insertCmd.ExecuteScalar();
                    }
                }
                catch (SqlException ex)
                {
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.WriteLine(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    tran.Rollback();
                    var e = new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                    throw e;
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
                _ => ""
            };
        }
        private int Dopasuj_Przyczyne(RodzajNieobecnosci rodzaj)
        {
            return rodzaj switch
            {
                RodzajNieobecnosci.ZL => 2,        // Zwolnienie lekarskie
                RodzajNieobecnosci.ZR => 3,        // Wypadek w pracy/choroba zawodowa
                RodzajNieobecnosci.ZY => 4,        // Wypadek w drodze do/z pracy
                RodzajNieobecnosci.ZZ => 5,        // Zwolnienie w okresie ciąży
                RodzajNieobecnosci.ZK => 9,        // Opieka nad dzieckiem do lat 14
                RodzajNieobecnosci.ZC => 10,       // Opieka nad inną osobą
                RodzajNieobecnosci.ZS => 11,       // Leczenie szpitalne
                RodzajNieobecnosci.UK => 12,       // Badanie dawcy/pobranie organów
                _ => 1                             // Nie dotyczy dla pozostałych przypadków
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
