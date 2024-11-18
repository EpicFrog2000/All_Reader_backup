using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace All_Readeer
{
    internal class Grafik_Pracy_Reader_v2024
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
        }
        private class Grafik
        {
            public string nazwa_pliku { get; set; } = "";
            public int nr_zakladki = 0;
            public int rok { get; set; } = 0;
            public int miesiac { get; set; } = 0;
            public List<Dane_Dni> dane_dni = [];
            public void Set_Miesiac(string wartosc)
            {
                var mies = wartosc.Trim().ToLower();
                if(mies == "pażdziernik")
                {
                    mies = "październik";
                }
                if (mies == "styczeń")
                {
                    miesiac = 1;
                }
                else if (mies == "luty")
                {
                    miesiac = 2;
                }
                else if (mies == "marzec")
                {
                    miesiac = 3;
                }
                else if (mies == "kwiecień")
                {
                    miesiac = 4;
                }
                else if (mies == "maj")
                {
                    miesiac = 5;
                }
                else if (mies == "czerwiec")
                {
                    miesiac = 6;
                }
                else if (mies == "lipiec")
                {
                    miesiac = 7;
                }
                else if (mies == "sierpień")
                {
                    miesiac = 8;
                }
                else if (mies == "wrzesień")
                {
                    miesiac = 9;
                }
                else if (mies == "październik")
                {
                    miesiac = 10;
                }
                else if (mies == "listopad")
                {
                    miesiac = 11;
                }
                else if (mies == "grudzień")
                {
                    miesiac = 12;
                }
                else
                {
                    miesiac = 0;
                }
            }
        }
        private class Dane_Dni
        {
            public Pracownik pracownik = new();
            public List<Dane_Dnia> dane_dnia = [];
        }
        private class Dane_Dnia
        {
            public int dzien = 0;
            public List<Godz_Pracy> godz_pracy = [];
        }
        private class Godz_Pracy
        {
            public TimeSpan Godz_Rozpoczecia_Pracy = TimeSpan.Zero;
            public TimeSpan Godz_Zakonczenia_Pracy = TimeSpan.Zero;
        }
        private class CurrentPosition
        {
            public int row { get; set; } = 1;
            public int col { get; set; } = 1;
        }
        private string Optima_Connection_String = "";
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time)
        {
            Grafik grafik = new();
            try
            {
                grafik.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                grafik.nr_zakladki = Program.error_logger.Nr_Zakladki;
                Get_Header_Karta_Info(worksheet, ref grafik);
                Get_Dane_Dni(worksheet, ref grafik);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }

            try
            {
                Wpierdol_Plan_do_Optimy(grafik);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            if (string.IsNullOrEmpty(NewConnectionString))
            {
                Program.error_logger.New_Custom_Error("Error: Empty Connection string in gv2024");
                Console.WriteLine("Error: Empty Connection string in gv2024");
                throw new Exception("Error: Empty Connection string in gv2024");
            }
            Optima_Connection_String = NewConnectionString;
        }
        private void Get_Header_Karta_Info(IXLWorksheet worksheet, ref Grafik grafik)
        {
            var dane = worksheet.Cell(1, 1).GetValue<string>().Trim();
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Tytułu Grafiku", 1, 1, "Brak Tytułu Grafiku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            else
            {
                grafik.Set_Miesiac(dane.Split(' ')[3].Trim());
                if (grafik.miesiac == 0)
                {
                    Program.error_logger.New_Error(dane, "Miesiac", 1, 1, "Źle wpisany miesiąc");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if (int.TryParse(dane.Split(' ')[5].Trim(), out int parsedYear))
                {
                    grafik.rok = parsedYear;
                }
                else
                {
                    Program.error_logger.New_Error(dane, "Rok", 1, 1, "Źle wpisany rok");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if (grafik.rok == 0)
                {
                    Program.error_logger.New_Error(dane, "Rok", 1, 1, "Źle wpisany rok");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
        }
        private void Get_Dane_Dni(IXLWorksheet worksheet, ref Grafik grafik)
        {
            CurrentPosition poz = new() { col = 1, row = 4 };
            while (true)
            {
                Dane_Dni dane_dni = new();
                //get pracownika
                var nazwiskoimie = worksheet.Cell(poz.row, poz.col).GetValue<string>().Trim();
                if (string.IsNullOrEmpty(nazwiskoimie))
                {
                    break;
                }
                try
                {
                    dane_dni.pracownik.Nazwisko = nazwiskoimie.Split(" ")[0];
                    dane_dni.pracownik.Imie = nazwiskoimie.Split(" ")[1];
                }
                catch (Exception ex)
                {
                    Program.error_logger.New_Error(nazwiskoimie, "Nazwisko Imie", poz.col, poz.row, "Źle wpisane nazwisko i imie " + ex.Message);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // get wysokosc wpisu dla osoby
                int height = 0;
                while (true)
                {
                    height++;
                    var dane = worksheet.Cell(poz.row + height, poz.col).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(dane))
                    {
                        break;
                    }
                }
                //get dane dni miesiaca
                int j = 0;
                for (int i = 0; i < 31; i++)
                {
                    Dane_Dnia dane_dnia = new();
                    //get dzien nr
                    var dziennr = worksheet.Cell(3, poz.col + 1 + j).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(dziennr))
                    {
                        if (int.TryParse(dziennr, out int parsedDzien))
                        {
                            dane_dnia.dzien = parsedDzien;
                        }
                        else if (DateTime.TryParse(dziennr, out DateTime Data))
                        {
                            dane_dnia.dzien = Data.Day;
                        }
                        else
                        {
                            Program.error_logger.New_Error(dziennr, "dzien", poz.col, 5, "Błędny nr dnia");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        if (dane_dnia.dzien > 31 || dane_dnia.dzien == 0)
                        {
                            break;
                        }
                        // get godziny pracy dnia
                        for (int k = 0; k < height; k++)
                        {
                            Godz_Pracy godziny = new();
                            var godzr = worksheet.Cell(poz.row + k, poz.col + 1 + j).GetValue<string>().Trim();
                            if (!string.IsNullOrEmpty(godzr) && godzr != "" && godzr.Length > 0)
                            {
                                try
                                {
                                    godziny.Godz_Rozpoczecia_Pracy = Reader.Try_Get_Date(godzr);
                                }
                                catch(Exception ex)
                                {
                                    Program.error_logger.New_Error(godzr, "Godz_Rozpoczecia_Pracy", poz.col + 1 + j, poz.row + k, ex.Message);
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }

                            }
                            var godzz = worksheet.Cell(poz.row + k, poz.col + 1 + j + 1).GetValue<string>().Trim();
                            if (!string.IsNullOrEmpty(godzz) && godzz != "" && godzz.Length > 0)
                            {
                                try
                                {
                                    godziny.Godz_Zakonczenia_Pracy = Reader.Try_Get_Date(godzz);
                                }
                                catch (Exception ex)
                                {
                                    Program.error_logger.New_Error(godzz, "Godz_Zakonczenia_Pracy", poz.col + 1 + j + 1, poz.row + k, ex.Message);
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }
                            }
                            if (godziny.Godz_Rozpoczecia_Pracy != TimeSpan.Zero && godziny.Godz_Zakonczenia_Pracy != TimeSpan.Zero)
                            {
                                dane_dnia.godz_pracy.Add(godziny);
                            }
                        }
                    }
                    j += 2;
                    dane_dni.dane_dnia.Add(dane_dnia);
                }
                grafik.dane_dni.Add(dane_dni);
                poz.row += height+1;
            }
        }
        private void Wpierdol_Plan_do_Optimy(Grafik grafik)
        {
            using (SqlConnection connection = new SqlConnection(Optima_Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                foreach (var dane_Dni in grafik.dane_dni)
                {
                    foreach (var dzien in dane_Dni.dane_dnia)
                    {
                        try
                        {
                            foreach (var godziny in dzien.godz_pracy)
                            {
                                // jak praca po północy to na next dzien przeniesc
                                if (godziny.Godz_Zakonczenia_Pracy < godziny.Godz_Rozpoczecia_Pracy)
                                {
                                    // insert godziny przed północą
                                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertPlanDoOptimy, connection, tran))
                                    {
                                        insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{grafik.rok}-{grafik.miesiac:D2}-{dzien.dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                        DateTime baseDate = new DateTime(1899, 12, 30);
                                        DateTime godzOdDate = baseDate + godziny.Godz_Rozpoczecia_Pracy;
                                        DateTime godzDoDate = baseDate + TimeSpan.FromHours(24);
                                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                        double czasPrzepracowanyInsert = (TimeSpan.FromHours(24) - godziny.Godz_Rozpoczecia_Pracy).TotalHours;
                                        insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                        insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                        insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", dane_Dni.pracownik.Nazwisko);
                                        insertCmd.Parameters.AddWithValue("@PracownikImieInsert", dane_Dni.pracownik.Imie);
                                        insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                                        insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                                        insertCmd.ExecuteScalar();
                                    }
                                    // insert po północy
                                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertPlanDoOptimy, connection, tran))
                                    {
                                        var data = new DateTime(grafik.rok, grafik.miesiac, dzien.dzien).AddDays(1);
                                        insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{data:yyyy-MM-dd}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                        DateTime baseDate = new DateTime(1899, 12, 30);
                                        DateTime godzOdDate = baseDate + TimeSpan.FromHours(0);
                                        DateTime godzDoDate = baseDate + godziny.Godz_Zakonczenia_Pracy;
                                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                        double czasPrzepracowanyInsert = godziny.Godz_Zakonczenia_Pracy.TotalHours;
                                        insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                        insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                        insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", dane_Dni.pracownik.Nazwisko);
                                        insertCmd.Parameters.AddWithValue("@PracownikImieInsert", dane_Dni.pracownik.Imie);
                                        insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                                        insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                                        insertCmd.ExecuteScalar();
                                    }
                                }
                                else
                                {
                                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertPlanDoOptimy, connection, tran))
                                    {
                                        insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{grafik.rok}-{grafik.miesiac:D2}-{dzien.dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                        insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = ("1899-12-30 " + godziny.Godz_Rozpoczecia_Pracy.ToString());
                                        insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = ("1899-12-30 " + godziny.Godz_Zakonczenia_Pracy.ToString());
                                        insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", (godziny.Godz_Zakonczenia_Pracy - godziny.Godz_Rozpoczecia_Pracy).TotalHours);
                                        insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", (godziny.Godz_Zakonczenia_Pracy - godziny.Godz_Rozpoczecia_Pracy).TotalHours);
                                        insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", dane_Dni.pracownik.Nazwisko);
                                        insertCmd.Parameters.AddWithValue("@PracownikImieInsert", dane_Dni.pracownik.Imie);
                                        insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                                        insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                                        insertCmd.ExecuteScalar();
                                    }
                                }
                            }
                        }
                        catch (SqlException ex)
                        {
                            tran.Rollback();
                            Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                            throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                        }
                        catch (FormatException)
                        {
                            continue;
                        }
                        catch (Exception ex)
                        {
                            tran.Rollback();
                            Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                            throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                        }
                    }
                }
                tran.Commit();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Poprawnie dodawno plan z pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
            }
        }
    }
}