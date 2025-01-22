using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace All_Readeer
{
    internal static class Grafik_Pracy_Reader_v2024
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
            public string Akronim { get; set; } = "";
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
        private class Current_Position
        {
            public int row { get; set; } = 1;
            public int col { get; set; } = 1;
        }
        public static void Process_Zakladka_For_Optima(IXLWorksheet worksheet)
        {
            try
            {
                var Lista_Pozycji_Grafików_Z_Zakladki = Find_Grafiki(worksheet);
                List<Grafik> grafiki = new();
                foreach (var pozycja in Lista_Pozycji_Grafików_Z_Zakladki)
                {
                    Grafik grafik = new();
                    grafik.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                    grafik.nr_zakladki = Program.error_logger.Nr_Zakladki;
                    Get_Header_Karta_Info(pozycja ,worksheet, ref grafik);
                    Get_Dane_Dni(pozycja, worksheet, ref grafik);
                    grafiki.Add(grafik);
                }
                if(grafiki.Count > 0)
                {
                    Dodaj_Plan_do_Optimy(grafiki);
                }
                
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        private static List<Current_Position> Find_Grafiki(IXLWorksheet worksheet)
        {
            List<Current_Position> Pozycje = new();
            int Limiter = 1000;
            int counter = 0;
            foreach (var cell in worksheet.CellsUsed())
            {

                try
                {
                    if (cell.HasFormula && !cell.Address.ToString()!.Equals(cell.FormulaA1))
                    {
                        counter++;
                        if(counter > Limiter)
                        {
                            break;
                        }
                        continue;
                    }

                    if (cell.Value.ToString().Contains("GRAFIK PRACY"))
                    {
                        Pozycje.Add(new Current_Position()
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
            return Pozycje;
        }
        private static void Get_Header_Karta_Info(Current_Position pozycja, IXLWorksheet worksheet, ref Grafik grafik)
        {
            var dane = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Tytułu Grafiku", pozycja.row, pozycja.col, "Brak Tytułu Grafiku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            else
            {
                grafik.Set_Miesiac(dane.Split(' ')[3].Trim());
                if (grafik.miesiac == 0)
                {
                    Program.error_logger.New_Error(dane, "Miesiac", pozycja.row, pozycja.col, "Źle wpisany miesiąc");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if (int.TryParse(dane.Split(' ')[5].Trim(), out int parsedYear))
                {
                    grafik.rok = parsedYear;
                }
                else
                {
                    Program.error_logger.New_Error(dane, "Rok", pozycja.row, pozycja.col, "Źle wpisany rok");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                if (grafik.rok == 0)
                {
                    Program.error_logger.New_Error(dane, "Rok", pozycja.row, pozycja.col, "Źle wpisany rok");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
        }
        private static void Get_Dane_Dni(Current_Position pozycja, IXLWorksheet worksheet, ref Grafik grafik)
        {
            pozycja.row += 3;
            while (true)
            {
                Dane_Dni dane_dni = new();
                //get pracownika
                var nazwiskoimie = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim().Replace("  ", " ");
                if (string.IsNullOrEmpty(nazwiskoimie))
                {
                    break;
                }
                try
                {
                    dane_dni.pracownik.Nazwisko = nazwiskoimie.Split(" ")[0];
                    dane_dni.pracownik.Imie = nazwiskoimie.Split(" ")[1];
                }
                catch
                {
                    //Program.error_logger.New_Error(nazwiskoimie, "Nazwisko Imie", pozycja.col, pozycja.row, "Źle wpisane nazwisko i imie. Wartość w komórce powinna być: Nazwisko Imie ");
                    //throw new Exception(Program.error_logger.Get_Error_String());
                }
                dane_dni.pracownik.Nazwisko = dane_dni.pracownik.Nazwisko.ToLower();
                dane_dni.pracownik.Nazwisko = char.ToUpper(dane_dni.pracownik.Nazwisko[0], CultureInfo.CurrentCulture) + dane_dni.pracownik.Nazwisko.Substring(1);
                dane_dni.pracownik.Imie = dane_dni.pracownik.Imie.ToLower();
                dane_dni.pracownik.Imie = char.ToUpper(dane_dni.pracownik.Imie[0], CultureInfo.CurrentCulture) + dane_dni.pracownik.Imie.Substring(1);
                // get wysokosc wpisu dla osoby
                int height = 0;
                while (true)
                {
                    height++;
                    var dane = worksheet.Cell(pozycja.row + height, pozycja.col).GetFormattedString().Trim();
                    if (!string.IsNullOrEmpty(dane))
                    {
                        break;
                    }
                }

                int j = 0;
                // try get akronim
                var dzienorakronim = worksheet.Cell(3, pozycja.col + 1).GetFormattedString().Trim();
                if(!string.IsNullOrEmpty(dzienorakronim) && dzienorakronim.ToLower().Contains("akronim"))
                {
                    dzienorakronim = worksheet.Cell(pozycja.row, pozycja.col + 1).GetFormattedString().Trim();
                    if (!string.IsNullOrEmpty(dzienorakronim))
                    {
                        dane_dni.pracownik.Akronim = dzienorakronim;
                    }
                    else
                    {
                        //TODO może error log brak akronimu
                    }
                    j++;
                }

                //get dane dni miesiaca
                for (int i = 0; i < 31; i++)
                {
                    Dane_Dnia dane_dnia = new();
                    //get dzien nr
                    var dziennr = worksheet.Cell(3, pozycja.col + 1 + j).GetFormattedString().Trim();
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
                            Program.error_logger.New_Error(dziennr, "dzien", pozycja.col, 5, "Błędny nr dnia");
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
                            var godzr = worksheet.Cell(pozycja.row + k, pozycja.col + 1 + j).GetFormattedString().Trim();
                            if (!string.IsNullOrEmpty(godzr) && godzr != "" && godzr.Length > 0)
                            {
                                try
                                {
                                    godziny.Godz_Rozpoczecia_Pracy = Reader.Try_Get_Date(godzr);
                                }
                                catch
                                {
                                    Program.error_logger.New_Error(godzr, "Godz Rozpoczecia Pracy", pozycja.col + 1 + j, pozycja.row + k, "Błędnie wpisany czas. Powinno być w formacie np. '08:00'");
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }

                            }
                            var godzz = worksheet.Cell(pozycja.row + k, pozycja.col + 1 + j + 1).GetFormattedString().Trim();
                            if (!string.IsNullOrEmpty(godzz) && godzz != "" && godzz.Length > 0)
                            {
                                try
                                {
                                    godziny.Godz_Zakonczenia_Pracy = Reader.Try_Get_Date(godzz);
                                }
                                catch
                                {
                                    Program.error_logger.New_Error(godzz, "Godz Zakonczenia Pracy", pozycja.col + 1 + j + 1, pozycja.row + k, "Błędnie wpisany czas. Powinno być w formacie np. '08:00'");
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
                pozycja.row += height+1;
            }
        }
        private static void Dodaj_Plan_do_Optimy(List<Grafik> grafiki)
        {
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    foreach (var grafik in grafiki)
                    {
                        foreach (var dane_Dni in grafik.dane_dni)
                        {
                            foreach (var dzien in dane_Dni.dane_dnia)
                            {
                                try
                                {
                                    foreach (var godziny in dzien.godz_pracy)
                                    {
                                        Zrob_Insert_Plan_command(connection, tran, grafik, dane_Dni.pracownik, new DateTime(grafik.rok, grafik.miesiac, dzien.dzien), godziny.Godz_Rozpoczecia_Pracy, godziny.Godz_Zakonczenia_Pracy);
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
                    }
                    tran.Commit();
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Poprawnie dodawno plan z pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
            }
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
        private static void Zrob_Insert_Plan_command(SqlConnection connection, SqlTransaction transaction, Grafik grafik, Pracownik pracownik, DateTime data, TimeSpan startGodz, TimeSpan endGodz)
        {
            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertPlanDoOptimy, connection, transaction))
            {
                DateTime baseDate = new DateTime(1899, 12, 30);
                DateTime godzOdDate = baseDate + startGodz;
                DateTime godzDoDate = baseDate + endGodz;
                insertCmd.Parameters.AddWithValue("@DataInsert", data);
                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                insertCmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(pracownik));
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
}