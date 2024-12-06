using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

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
        public static void Process_Zakladka_For_Optima(IXLWorksheet worksheet)
        {
            try
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
                        try
                        {
                            grafik.Set_Miesiac(worksheet.Name.Split(' ')[0]);
                            grafik.Rok = int.Parse(worksheet.Name.Split(' ')[1]);
                        }
                        catch
                        {
                            Program.error_logger.New_Custom_Error($"Zła nazwa zakładki. Ma wyglądać: miesiąc rok a jest {worksheet.Name}");
                            throw;
                        }
                        grafik.Pracownik = Get_Pracownik(worksheet, new Current_Position { row = Startpozycja.row - 2, col = Startpozycja.col + ((counter * 3) + 1) });
                        if (string.IsNullOrEmpty(grafik.Pracownik.Imie) && string.IsNullOrEmpty(grafik.Pracownik.Nazwisko) && string.IsNullOrEmpty(grafik.Pracownik.Akronim))
                        {
                            break;
                        }
                        var dane = Get_Dane_Dni(worksheet, new Current_Position { row = Startpozycja.row + 4, col = Startpozycja.col + ((counter * 3) + 1) });
                        foreach (var d in dane)
                        {
                            grafik.Dane_Dni.Add(d);
                        }
                        grafiki.Add(grafik);
                        counter++;
                    }
                    if (grafiki.Count > 0)
                    {
                        Dodaj_Plan_do_Optimy(grafiki);
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
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
                    Program.error_logger.New_Error(nazwiskoimie, "Nazwisko i Imie", pozycja.col, pozycja.row, "Zły format wpisanego nazwiska i imienia pracownika");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }

            var akronim = "";
            for (int i = 0; i < 3; i++)
            {
                if (string.IsNullOrEmpty(akronim))
                {
                    akronim = worksheet.Cell(pozycja.row+1, pozycja.col + i).GetFormattedString().Trim();
                }
            }
            pracownik.Akronim = akronim;
            if (!string.IsNullOrEmpty(nazwiskoimie))
            {
                pracownik.Nazwisko = nazwiskoimie.Split(" ")[0].Trim();
                pracownik.Imie = nazwiskoimie.Split(" ")[1].Trim();
            }
            return pracownik;
        }
        private static List<Dane_Dnia> Get_Dane_Dni(IXLWorksheet worksheet, Current_Position pozycja)
        {
            List<Dane_Dnia> Dane_Dni = new();
            for (int i = 0; i < 31; i++)
            {
                var dane = "";
                try
                {
                    Dane_Dnia dane_Dnia = new Dane_Dnia();
                    dane_Dnia.Nr_Dnia = i + 1;
                    dane = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
                    if (string.IsNullOrEmpty(dane))
                    {
                        pozycja.row += 1;
                        continue;
                    }
                    dane_Dnia.Godzina_Pracy_Od = TimeSpan.Parse(dane);
                    dane = worksheet.Cell(pozycja.row, pozycja.col + 1).GetFormattedString().Trim();
                    if (string.IsNullOrEmpty(dane))
                    {
                        pozycja.row += 1;
                        continue;
                    }
                    dane_Dnia.Godzina_Pracy_Do = TimeSpan.Parse(dane);
                    Dane_Dni.Add(dane_Dnia);
                    pozycja.row += 1;
                }
                catch
                {
                    Program.error_logger.New_Error(dane, "Godzina", pozycja.col, pozycja.row, "Błędna wartość w godzinie");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
            return Dane_Dni;
        }
        private static void Dodaj_Plan_do_Optimy(List<Grafik> grafiki)
        {
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                foreach (var grafik in grafiki)
                {
                    foreach (var dane_DniA in grafik.Dane_Dni)
                    {
                        try
                        {
                            Zrob_Insert_Plan_command(connection, tran, grafik, grafik.Pracownik, DateTime.ParseExact($"{grafik.Rok}-{grafik.Miesiac:D2}-{dane_DniA.Nr_Dnia:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), dane_DniA.Godzina_Pracy_Od, dane_DniA.Godzina_Pracy_Do);
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
                tran.Commit();
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