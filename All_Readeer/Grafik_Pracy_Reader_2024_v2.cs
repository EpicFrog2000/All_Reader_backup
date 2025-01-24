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
                if (string.IsNullOrEmpty(wartosc))
                {
                    Miesiac = -1;
                    return;
                }
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
            foreach (IXLCell? cell in worksheet.CellsUsed())
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
                List<Current_Position> Lista_Pozycji_Grafików_Z_Zakladki = Find_Grafiki(worksheet);
                List<Grafik> grafiki = new();
                foreach (Current_Position Startpozycja in Lista_Pozycji_Grafików_Z_Zakladki)
                {
                    Current_Position pozycja = Startpozycja;
                    int counter = 0;
                    while (true)
                    {
                        int rowOffset = -5;
                        string dane = "";
                        Grafik grafik = new();
                        try
                        {
                            dane = worksheet.Cell(pozycja.row + rowOffset, pozycja.col + 3).GetFormattedString().Trim();
                        }
                        catch
                        {
                            rowOffset = -4;
                            try
                            {
                                dane = worksheet.Cell(pozycja.row + rowOffset, pozycja.col + 3).GetFormattedString().Trim();
                            }
                            catch
                            {
                                rowOffset = -3;
                                try
                                {
                                    dane = worksheet.Cell(pozycja.row + rowOffset, pozycja.col + 3).GetFormattedString().Trim();
                                }
                                catch
                                {
                                    Program.error_logger.New_Error(dane, "Naglowek", pozycja.col + 3, pozycja.row - 5, "Zły format pliku");
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }
                            }
                        }// xddddddddddd
                        grafik.Set_Miesiac(dane);
                        if(grafik.Miesiac < 1)
                        {
                            Program.error_logger.New_Error(dane, "Miesiac", pozycja.col + 3, pozycja.row + rowOffset, "Błędna wartość w mieisac");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }


                        dane = worksheet.Cell(pozycja.row + rowOffset, pozycja.col + 6).GetFormattedString().Trim();
                        if (string.IsNullOrEmpty(dane))
                        {
                            Program.error_logger.New_Error(dane, "Rok", pozycja.col + 5, pozycja.row + rowOffset, "Błędna wartość w rok");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }

                        if (int.TryParse(dane, out int tmprok))
                        {
                            grafik.Rok = tmprok;
                        }
                        else
                        {
                            Program.error_logger.New_Error(dane, "Rok", pozycja.col + 5, pozycja.row + rowOffset, "Błędna wartość w rok");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        
                        grafik.Pracownik = Get_Pracownik(worksheet, new Current_Position { row = Startpozycja.row, col = Startpozycja.col + ((counter * 3) + 1) });
                        if (string.IsNullOrEmpty(grafik.Pracownik.Imie) && string.IsNullOrEmpty(grafik.Pracownik.Nazwisko) && string.IsNullOrEmpty(grafik.Pracownik.Akronim))
                        {
                            break;
                        }

                        List<Dane_Dnia> dane2 = Get_Dane_Dni(worksheet, new Current_Position { row = Startpozycja.row + 4, col = Startpozycja.col + ((counter * 3) + 1) });
                        foreach(Dane_Dnia d in dane2)
                        {
                            grafik.Dane_Dni.Add(d);
                        }
                        grafiki.Add(grafik);
                        counter++;
                    }
                }
                if (grafiki.Count > 0)
                {
                    Dodaj_Plan_do_Optimy(grafiki);
                }
                else
                {
                    Program.error_logger.New_Custom_Error("Zły format pliku, nie znaleniono żadnych grafików z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception("Zły format pliku, nie znaleniono żadnych grafików z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }

        private static Pracownik Get_Pracownik(IXLWorksheet worksheet, Current_Position pozycja)
        {
            Pracownik pracownik = new Pracownik();
            string pole1 = "";
            string pole2 = "";
            int offset = 0;
            while (true)
            {
                pozycja.row--;
                if(pozycja.row < 1)
                {
                    return pracownik;
                }
                pole1 = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
                if (pole1 != "Godziny pracy od")
                {
                    offset++;
                    for (int i = 0; i < 3; i++)
                    {
                        pole1 = worksheet.Cell(pozycja.row, pozycja.col+i).GetFormattedString().Trim();
                        if (!string.IsNullOrEmpty(pole1))
                        {
                            if(offset == 1)
                            {
                                pole2 = worksheet.Cell(pozycja.row-1, pozycja.col).GetFormattedString().Trim();
                            }
                            if(offset == 2)
                            {
                                pole2 = pole1;
                            }
                            break;
                        }
                    }
                    if (!string.IsNullOrEmpty(pole2))
                    {
                        break;
                    }
                }
            }
            if (!string.IsNullOrEmpty(pole1) && int.TryParse(pole1, out int impa))
            {
                pracownik.Akronim = pole1;
                if (!string.IsNullOrEmpty(pole2))
                {
                    string[] parts = pole2.Split(" ");
                    if(parts.Length == 2)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                    }
                }
            }
            else
            {
                if (!string.IsNullOrEmpty(pole2))
                {
                    string[] parts = pole2.Split(" ");
                    if (parts.Length == 2)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                    }
                    else if (parts.Length == 3)
                    {
                        pracownik.Nazwisko = parts[0];
                        pracownik.Imie = parts[1];
                        if (int.TryParse(parts[2], out int tmpint))
                        {
                            pracownik.Akronim = parts[2];
                        }
                    }
                    else
                    {
                        Program.error_logger.New_Error(pole1, "Imie nazwisko akronim", pozycja.col, pozycja.row, "Błędny format danych w komórkach imie nazwisko akronim");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
            }
            return pracownik;
        }
        private static List<Dane_Dnia> Get_Dane_Dni(IXLWorksheet worksheet, Current_Position pozycja)
        {
            List<Dane_Dnia> Dane_Dni = new();
            for (int i = 0; i < 31; i++)
            {
                string dane = "";
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
                    if (TimeSpan.TryParse(dane, out TimeSpan time))
                    {
                        dane_Dnia.Godzina_Pracy_Od = time;
                    }
                    else
                    {
                        Program.error_logger.New_Error(dane, "Godzina pracy od", pozycja.col, pozycja.row, "Błędna wartość w godzinie pracy od");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }

                    dane = worksheet.Cell(pozycja.row, pozycja.col + 1).GetFormattedString().Trim();
                    if (string.IsNullOrEmpty(dane))
                    {
                        pozycja.row += 1;
                        continue;
                    }
                    if (TimeSpan.TryParse(dane, out TimeSpan time2))
                    {
                        dane_Dnia.Godzina_Pracy_Do = time2;
                    }
                    else
                    {
                        Program.error_logger.New_Error(dane, "Godzina pracy do", pozycja.col + 1, pozycja.row, "Błędna wartość w godzinie pracy do");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
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
            int dodano = 0;
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    foreach (Grafik grafik in grafiki)
                    {
                        foreach (Dane_Dnia dane_DniA in grafik.Dane_Dni)
                        {
                            try
                            {
                                dodano += Zrob_Insert_Plan_command(connection, tran, grafik, grafik.Pracownik, DateTime.ParseExact($"{grafik.Rok}-{grafik.Miesiac:D2}-{dane_DniA.Nr_Dnia:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture), dane_DniA.Godzina_Pracy_Od, dane_DniA.Godzina_Pracy_Do);
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
                }
                if(dodano > 0)
                {
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno plan z pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }
        }
        private static int Get_ID_Pracownika(Pracownik pracownik)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
                {
                    using (SqlCommand getCmd = new SqlCommand(Program.sqlQueryGetPRI_PraId, connection))
                    {
                        connection.Open();
                        getCmd.Parameters.AddWithValue("@Akronim", pracownik.Akronim);
                        getCmd.Parameters.AddWithValue("@PracownikImieInsert", pracownik.Imie);
                        getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", pracownik.Nazwisko);
                        object result = getCmd.ExecuteScalar();
                        return result != null ? Convert.ToInt32(result) : 0;
                    }
                }

            }
            catch (Exception ex)
            {
                Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakładki: " + Program.error_logger.Nr_Zakladki + " nazwa zakładki: " + Program.error_logger.Nazwa_Zakladki);
                throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakładki {Program.error_logger.Nr_Zakladki}" + " nazwa zakładki: " + Program.error_logger.Nazwa_Zakladki);
            }
        }

        private static int Zrob_Insert_Plan_command(SqlConnection connection, SqlTransaction transaction, Grafik grafik, Pracownik pracownik, DateTime data, TimeSpan startGodz, TimeSpan endGodz)
        {
            int IdPracowkika = Get_ID_Pracownika(pracownik);
            using (SqlCommand cmd = new(@"
IF EXISTS (
SELECT 1 
FROM cdn.PracPlanDni 
WHERE PPL_Data = @DataInsert 
    AND PPL_PraId = @PRI_PraId
)
BEGIN
IF EXISTS (
    SELECT 1 
    FROM cdn.PracPlanDniGodz 
    WHERE PGL_PplId = (
        SELECT PPL_PplId 
        FROM cdn.PracPlanDni 
        WHERE PPL_Data = @DataInsert 
            AND PPL_PraId = @PRI_PraId
    )
        AND PGL_OdGodziny = @GodzOdDate 
        AND PGL_DoGodziny = @GodzDoDate
)
BEGIN
    SELECT 1;
END
ELSE
BEGIN
    SELECT 0;
END
END
ELSE
BEGIN
SELECT 0;
END", connection, transaction))
            {
                cmd.Parameters.AddWithValue("@DataInsert", data);
                cmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(Program.baseDate + startGodz);
                cmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(Program.baseDate + endGodz);
                cmd.Parameters.AddWithValue("@PRI_PraId", IdPracowkika);
                if ((int)cmd.ExecuteScalar() == 1)
                {
                    return 0;
                }
            }
            using (SqlCommand insertCmd = new(Program.sqlQueryInsertPlanDoOptimy, connection, transaction))
            {
                insertCmd.Parameters.AddWithValue("@DataInsert", data);
                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = (DateTime)(Program.baseDate + startGodz);
                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = (DateTime)(Program.baseDate + endGodz);
                insertCmd.Parameters.AddWithValue("@PRI_PraId", IdPracowkika);
                insertCmd.Parameters.AddWithValue("@ImieMod", Truncate(Program.error_logger.Last_Mod_Osoba, 20));
                insertCmd.Parameters.AddWithValue("@NazwiskoMod", Truncate(Program.error_logger.Last_Mod_Osoba, 50));
                insertCmd.Parameters.AddWithValue("@DataMod", Program.error_logger.Last_Mod_Time);

                insertCmd.ExecuteScalar();
            }
            return 1;
        }

        private static string Truncate(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value))
            {
                return string.Empty;
            }
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }
    }

}
