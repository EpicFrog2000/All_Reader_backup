using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Data.SqlClient;

namespace All_Readeer
{
    internal class Konduktor_Reader
    {
        private class Konduktor
        {
            public int Nr_Sluzbowy { get; set; }
            public string Stanowisko { get; set; } = "";
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
        }
        private class Karta_Ewidencji
        {
            public Konduktor Konduktor { get; set; } = new();
            public List<Karta_Ewidencji_Detale> karta_Ewidencji_Detale { get; set; } = new();
            public int Miesiac { get; set; }
            public int Rok { get; set; }
            public int Nominal_M_CA { get; set; }
            public decimal Total_Godziny_Lacznie_Z_Odpoczynkiem { get; set; }
            public decimal Total_Godziny_Pracy { get; set; }
            public decimal Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem { get; set; }
            public decimal Total_Liczba_Godzin_Relacji_Pracy { get; set; }
            public decimal Total_Liczba_Godzin_Nocnych { get; set; }
            public decimal Total_Liczba_NadGodzin_Ogolem_50 { get; set; }
            public decimal Total_Liczba_NadGodzin_Ogolem_100 { get; set; }
            public decimal Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 { get; set; }
            public decimal Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 { get; set; }
            public decimal Suma_Fodzin_Przepracowanych_Plus_Absencja { get; set; }
            public decimal Nadgodziny_50 { get; set; }
            public decimal Nadgodziny_100 { get; set; }
            public decimal Total_Liczba_Godzin_Absencji { get; set; }
            public decimal Total_Sprzedaz_Bilety_Zagranica { get; set; }
            public decimal Total_Sprzedaz_Bilety_Kraj { get; set; }
            public decimal Total_Sprzedaz_Bilety_Globalne { get; set; }
            public decimal Total_Sprzedaz_Bilety_Wartosc_Towarow { get; set; }
            public decimal Total_Liczba_Napojow_Awaryjnych { get; set; }
            public void Set_Miesiac(string nazwa)
            {
                if (!string.IsNullOrEmpty(nazwa))
                {
                    if (nazwa.ToLower() == "styczeń")
                    {
                        Miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "luty")
                    {
                        Miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "marzec")
                    {
                        Miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "kwiecień")
                    {
                        Miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "maj")
                    {
                        Miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "czerwiec")
                    {
                        Miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "lipiec")
                    {
                        Miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "sierpień")
                    {
                        Miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "wrzesień")
                    {
                        Miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "październik")
                    {
                        Miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "listopad")
                    {
                        Miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "grudzień")
                    {
                        Miesiac = 12;
                    }
                    else
                    {
                        Miesiac = 0;
                    }
                }

            }
        }
        private class Karta_Ewidencji_Detale
        {
            public List<Detale_Dzien> Lista_Detali_Dni = new();
            public string Numer_Relacji { get; set; } = "";
            public string Relacja { get; set; } = "";

            public string Nr_Pociagu { get; set; } = "";
            public double Liczba_Godzin_Relacji_Z_Odpoczynkiem { get; set; }
            public double Liczba_Godzin_Relacji_Pracy { get; set; }
            public double Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 { get; set; }
            public double Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 { get; set; }
            public double Wartosc_Biletow_Zagranica { get; set; }
            public double Wartosc_Biletow_Kraj { get; set; }
            public double Wartosc_Biletow_Globalne { get; set; }
            public double Wartosc_Towarow { get; set; }
        }
        private class Detale_Dzien
        {
            public int Dzien { get; set; }
            public List<TimeSpan> Godziny_Pracy_Od { get; set; } = new();
            public List<TimeSpan> Godziny_Pracy_Do { get; set; } = new();
            public List<TimeSpan> Godziny_Odpoczynku_Od { get; set; } = new();
            public List<TimeSpan> Godziny_Odpoczynku_Do { get; set; } = new();
            public int Godziny_Lacznie_Z_Odpoczynkiem { get; set; }
            public int Godziny_Pracy { get; set; }
            public int Liczba_Godzin_Nocnych { get; set; }
            public int Liczba_NadGodzin_Ogolem_50 { get; set; }
            public int Liczba_NadGodzin_Ogolem_100 { get; set; }
            public string Nazwa_Absencji { get; set; } = "";
            public int Liczba_Godzin_Absencji { get; set; }
            public int Liczba_Napojow_Awaryjnych { get; set; }
        }
        private class ErrorLogger
        {
            public string Nazwa_Pliku = "";
            public int Nr_Zakladki = 0;
            public string FilePath = "";
            private string Wartosc_Pola = "";
            private string Nazwa_Pola = "";
            private int Kolumna = -1;
            private int Rzad = -1;
            private DateTime Data_Czas_Wykrycia_Bledu;

            public void New_Error(string wartoscPola, string nazwaPola, int kolumna, int rzad)
            {
                Nazwa_Pola = nazwaPola;
                Wartosc_Pola = wartoscPola;
                Kolumna = kolumna;
                Rzad = rzad;
                Data_Czas_Wykrycia_Bledu = DateTime.Now;
                Append_Error_To_File();
            }

            public string Get_Error_String()
            {
                return $"W pliku {Nazwa_Pliku} wystąpił błąd w zakładce nr {Nr_Zakladki}: W Kolumnie {Kolumna}, rzędzie {Rzad}, Powinna znaleźć się wartość {Nazwa_Pola} a jest: \"{Wartosc_Pola}\". Data wykrycia: {Data_Czas_Wykrycia_Bledu}";
            }

            public void Print_Outside_Error(string Error_Msg)
            {
                Console.WriteLine(Error_Msg);
            }

            public void Append_Error_To_File()
            {
                var ErrorsLogFile = Path.Combine(FilePath, "Errors.txt");
                if (!File.Exists(ErrorsLogFile))
                {
                    File.Create(ErrorsLogFile).Dispose();
                }
                File.AppendAllText(ErrorsLogFile, Get_Error_String() + Environment.NewLine);
            }


        }
        private static string FilePath = "";
        private static string Connection_String = "Server=ITEGER-NT;Database=nowaBazaZadanie;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
        private static Karta_Ewidencji karta_Ewidencji = new();
        private static IXLWorksheet worksheet = null!;
        private static List<string> Errors = new();
        private static ErrorLogger Error_Logger = new(); // TODO: <- jest tylko zadeklarowany, dopisz wykrywanie błędów w trakcie czytania pliku itp
        private static DateTime Last_Mod_Time = DateTime.Now;
        private static string Last_Mod_Osoba = "";
        public static void Set_Errors_File_Folder(string NewfilePath)
        {
            Error_Logger.FilePath = NewfilePath;
        }
        public static List<string> Process_To_Db()
        {
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            using (var workbook = new XLWorkbook(FilePath))
            {
                try
                {
                    // TODO zrob żeby dzialalo na wszystkie karty
                    worksheet = workbook.Worksheet(1);
                    Get_Ogolne_Data();
                    Get_Konduktor();
                    Get_Data_Relacje();
                    Insert_Konduktorzy_To_Db([karta_Ewidencji.Konduktor]);
                    Insert_Karta_Ewidencji_To_Db(karta_Ewidencji);
                }
                catch(Exception ex)
                {
                    Errors.Add(DateTime.Now.ToString() + " |Konduktor_Reader|Program Error: " + ex.ToString());
                }
            }
            return Errors;
        }
        public static void Set_File_Path(string NewfilePath)
        {
            FilePath = NewfilePath;
        }
        public static void Set_Db_Tables_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }

        private static (string, DateTime) Get_File_Meta_Info()
        {
            try
            {
                using (var workbook = new XLWorkbook(FilePath))
                {
                    string lastModifiedBy = workbook.Properties.LastModifiedBy!;
                    DateTime lastWriteTime = File.GetLastWriteTime(FilePath);
                    return (lastModifiedBy, lastWriteTime);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                return ("Error", DateTime.Now);
            }
        }
        private static void Get_Konduktor()
        {
            Konduktor konduktor = new();
            var strnumer = worksheet.Cell(2, 23).GetFormattedString().Trim();
            if (!string.IsNullOrEmpty(strnumer))
            {
                var s = strnumer.Split(' ');
                konduktor.Imie = s[0];
                konduktor.Nazwisko = s[1];
            }
            strnumer = worksheet.Cell(3, 23).GetFormattedString().Trim();
            konduktor.Nr_Sluzbowy = TrySetNum(strnumer);
            strnumer = worksheet.Cell(3, 1).GetFormattedString().Trim();
            if (!string.IsNullOrEmpty(strnumer))
            {
                var s = strnumer.Split(':');
                konduktor.Stanowisko = s[1];
            }
            karta_Ewidencji.Konduktor = konduktor;
        }
        private static void Get_Ogolne_Data()
        {
            var strnumer = worksheet.Cell(1, 1).GetFormattedString().Trim();
            while (strnumer.Contains("  "))
            {
                strnumer = strnumer.Replace("  ", " ");
            }

            if (!string.IsNullOrEmpty(strnumer))
            {
                var s = strnumer.Split(' ');
                karta_Ewidencji.Set_Miesiac(s[7]);
                karta_Ewidencji.Rok = TrySetNum(s[9]);
            }

            strnumer = worksheet.Cell(2, 14).GetFormattedString().Trim();
            karta_Ewidencji.Nominal_M_CA = TrySetNum(strnumer);

            strnumer = worksheet.Cell(40, 9).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 10).GetFormattedString().Trim();
            karta_Ewidencji.Total_Godziny_Pracy = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 11).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 12).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Relacji_Pracy = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 13).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Nocnych = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 14).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_NadGodzin_Ogolem_50 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 15).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_NadGodzin_Ogolem_100 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 16).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 17).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 19).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Godzin_Absencji = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 20).GetFormattedString().Trim();
            karta_Ewidencji.Total_Sprzedaz_Bilety_Zagranica = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 21).GetFormattedString().Trim();
            karta_Ewidencji.Total_Sprzedaz_Bilety_Kraj = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 22).GetFormattedString().Trim();
            karta_Ewidencji.Total_Sprzedaz_Bilety_Globalne = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 23).GetFormattedString().Trim();
            karta_Ewidencji.Total_Sprzedaz_Bilety_Wartosc_Towarow = TrySetNum(strnumer);
            strnumer = worksheet.Cell(40, 24).GetFormattedString().Trim();
            karta_Ewidencji.Total_Liczba_Napojow_Awaryjnych = TrySetNum(strnumer);
            strnumer = worksheet.Cell(41, 10).GetFormattedString().Trim();
            karta_Ewidencji.Suma_Fodzin_Przepracowanych_Plus_Absencja = TrySetNum(strnumer);
            strnumer = worksheet.Cell(41, 11).GetFormattedString().Trim();
            karta_Ewidencji.Nadgodziny_50 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(41, 12).GetFormattedString().Trim();
            karta_Ewidencji.Nadgodziny_100 = TrySetNum(strnumer);
        }
        private static void Get_Data_Relacje()
        {
            int startRow = 8;
            List<Detale_Dzien> listadni = new();
            List<int> listaIndexow = Get_Indexy(startRow);
            int maxdzien = Get_Max_Dzien(startRow);
            for (int i = 0; i < listaIndexow.Count - 1; i++)
            {
                Karta_Ewidencji_Detale Relacja = new();
                Get_Dane_Relacji(ref Relacja, startRow - 1 + listaIndexow[i], startRow - 1 + listaIndexow[i + 1]);
                Get_Dane_Dni(ref Relacja, startRow - 1 + listaIndexow[i], startRow - 1 + listaIndexow[i + 1]);
                karta_Ewidencji.karta_Ewidencji_Detale.Add(Relacja);
            }
            Karta_Ewidencji_Detale Relacja2 = new();
            Get_Dane_Relacji(ref Relacja2, startRow - 1 + listaIndexow[listaIndexow.Count - 1], startRow - 1 + maxdzien);
            Get_Dane_Dni(ref Relacja2, startRow - 1 + listaIndexow[listaIndexow.Count - 1], startRow - 1 + maxdzien);
            karta_Ewidencji.karta_Ewidencji_Detale.Add(Relacja2);
        }
        private static void Get_Dane_Dni(ref Karta_Ewidencji_Detale Relacja, int startRelacji, int koniecRelacji)
        {
            for (int i = 0; i < koniecRelacji; i++)
            {
                Detale_Dzien dzien = new();
                var strnumer = worksheet.Cell(startRelacji + i, 1).GetFormattedString().Trim();
                dzien.Dzien = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 5).GetFormattedString().Trim();
                dzien.Godziny_Pracy_Od.Add(TryInsertTimeSpan(strnumer));
                strnumer = worksheet.Cell(startRelacji + i, 6).GetFormattedString().Trim();
                dzien.Godziny_Pracy_Do.Add(TryInsertTimeSpan(strnumer));
                strnumer = worksheet.Cell(startRelacji + i, 7).GetFormattedString().Trim();
                dzien.Godziny_Odpoczynku_Od.Add(TryInsertTimeSpan(strnumer));
                strnumer = worksheet.Cell(startRelacji + i, 8).GetFormattedString().Trim();
                dzien.Godziny_Odpoczynku_Do.Add(TryInsertTimeSpan(strnumer));
                strnumer = worksheet.Cell(startRelacji + i, 11).GetFormattedString().Trim();
                dzien.Godziny_Lacznie_Z_Odpoczynkiem = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 12).GetFormattedString().Trim();
                dzien.Godziny_Pracy = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 13).GetFormattedString().Trim();
                dzien.Liczba_Godzin_Nocnych = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 14).GetFormattedString().Trim();
                dzien.Liczba_NadGodzin_Ogolem_50 = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 15).GetFormattedString().Trim();
                dzien.Liczba_NadGodzin_Ogolem_100 = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 18).GetFormattedString().Trim();
                dzien.Nazwa_Absencji = strnumer;
                strnumer = worksheet.Cell(startRelacji + i, 19).GetFormattedString().Trim();
                dzien.Liczba_Godzin_Absencji = TrySetNum(strnumer);
                strnumer = worksheet.Cell(startRelacji + i, 25).GetFormattedString().Trim();
                dzien.Liczba_Napojow_Awaryjnych = TrySetNum(strnumer);
                Relacja.Lista_Detali_Dni.Add(dzien);
            }
        }
        private static TimeSpan TryInsertTimeSpan(string strnumer)
        {
            if (!string.IsNullOrEmpty(strnumer))
            {
                string[] tmps = strnumer.Contains(' ') ? strnumer.Split(' ') : [strnumer];

                foreach (var tmp in tmps)
                {
                    if (TimeSpan.TryParse(tmp, out TimeSpan tmpss))
                    {
                        return tmpss;
                    }
                }
            }

            return TimeSpan.Zero;
        }
        private static void Get_Dane_Relacji(ref Karta_Ewidencji_Detale Relacja, int startRelacji, int koniecRelacji)
        {
            var Rel = "";
            var strnumer = "";
            for (int i = 0; i < koniecRelacji - 1; i++)
            {
                strnumer = worksheet.Cell(startRelacji + i, 3).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(strnumer))
                {
                    break;
                }
                else
                {
                    if (strnumer != Rel)
                    {
                        Rel += strnumer;
                    }
                }
            }
            Relacja.Relacja = Rel.Trim().Split(' ')[0];
            Relacja.Nr_Pociagu = Rel.Trim().Split(' ')[1];
            strnumer = worksheet.Cell(startRelacji, 9).GetFormattedString().Trim();
            Relacja.Liczba_Godzin_Relacji_Z_Odpoczynkiem = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 10).GetFormattedString().Trim();
            Relacja.Liczba_Godzin_Relacji_Pracy = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 16).GetFormattedString().Trim();
            Relacja.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 17).GetFormattedString().Trim();
            Relacja.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 20).GetFormattedString().Trim();
            Relacja.Wartosc_Biletow_Zagranica = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 21).GetFormattedString().Trim();
            Relacja.Wartosc_Biletow_Kraj = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 22).GetFormattedString().Trim();
            Relacja.Wartosc_Biletow_Globalne = TrySetNum(strnumer);
            strnumer = worksheet.Cell(startRelacji, 23).GetFormattedString().Trim();
            Relacja.Wartosc_Towarow = TrySetNum(strnumer);
        }
        private static int Get_Max_Dzien(int startRow)
        {
            int max = TrySetNum(worksheet.Cell(startRow, 1).GetFormattedString().Trim());
            int counter = 1;
            while (true)
            {
                var strnumer = worksheet.Cell(startRow + counter, 1).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(strnumer))
                {
                    return max;
                }
                max = TrySetNum(strnumer);
                counter++;
            }
        }
        private static List<int> Get_Indexy(int startRow)
        {
            List<int> listaIndexow2 = new();
            while (true)
            {
                Karta_Ewidencji_Detale Relacja = new();
                var strnumer = worksheet.Cell(startRow, 1).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(strnumer))
                {
                    break;
                }
                if (!string.IsNullOrEmpty(worksheet.Cell(startRow, 2).GetFormattedString().Trim()))
                {
                    listaIndexow2.Add(TrySetNum(strnumer));
                }
                var tmpn = Skip_Do_Next_Relacji(startRow);
                if (tmpn == -1)
                {
                    break;
                }
                else
                {
                    startRow += tmpn;
                }
            }
            return listaIndexow2;
        }
        private static int Skip_Do_Next_Relacji(int currentRow)
        {
            var currentdzien = worksheet.Cell(currentRow, 1).GetFormattedString().Trim();
            for (int i = 1; TrySetNum(currentdzien) <= 31; i++)
            {
                var trygetlernum = worksheet.Cell(currentRow + i, 2).GetFormattedString().Trim();
                if (!string.IsNullOrEmpty(trygetlernum))
                {
                    return i;
                }

                currentdzien = worksheet.Cell(currentRow + i, 1).GetFormattedString().Trim();
                if (string.IsNullOrEmpty(currentdzien))
                {
                    break;
                }
            }
            return -1;
        }
        private static int TrySetNum(string strnumer)
        {
            var number = 0;
            if (!string.IsNullOrEmpty(strnumer))
            {
                if (int.TryParse(strnumer, out number))
                {
                    return number;
                }
                else
                {
                    return 0;
                }
            }
            else
            {
                return 0;
            }
        }
        private static void Insert_Konduktorzy_To_Db(List<Konduktor> konduktorzy)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    foreach (var konduktor in konduktorzy)
                    {
                        string checkQuery = "SELECT COUNT(1) FROM Karta_Ewidecji_Konduktorzy WHERE Imie = @Imie AND Nazwisko = @Nazwisko AND Nr_Sluzbowy = @Nr_Sluzbowy AND Stanowisko = @Stanowisko;";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection, tran))
                        {
                            checkCmd.Parameters.AddWithValue("@Imie", konduktor.Imie);
                            checkCmd.Parameters.AddWithValue("@Nazwisko", konduktor.Nazwisko);
                            checkCmd.Parameters.AddWithValue("@Nr_Sluzbowy", konduktor.Nr_Sluzbowy);
                            checkCmd.Parameters.AddWithValue("@Stanowisko", konduktor.Stanowisko);

                            int count = (int)checkCmd.ExecuteScalar();

                            if (count == 0)
                            {
                                string insertQuery = "INSERT INTO Karta_Ewidecji_Konduktorzy (Imie, Nazwisko, Nr_Sluzbowy, Stanowisko, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) VALUES (@Imie, @Nazwisko, @Nr_Sluzbowy, @Stanowisko, @Ostatnia_Modyfikacja_Data, @Ostatnia_Modyfikacja_Osoba)";
                                using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                                {
                                    insertCmd.Parameters.AddWithValue("@Imie", konduktor.Imie);
                                    insertCmd.Parameters.AddWithValue("@Nazwisko", konduktor.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@Nr_Sluzbowy", konduktor.Nr_Sluzbowy);
                                    insertCmd.Parameters.AddWithValue("@Stanowisko", konduktor.Stanowisko);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                                    insertCmd.ExecuteNonQuery();
                                }
                            }
                        }
                    }
                    tran.Commit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Errors.Add(DateTime.Now.ToString() + " |Konduktor_Reader|Insert to db Error: " + ex.ToString());
                    tran.Rollback();
                }
            }
        }
        private static void Insert_Karta_Ewidencji_To_Db(Karta_Ewidencji karta)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    var Id_Konduktora = 0;
                    string selectQuery = "SELECT Id_Konduktora FROM Karta_Ewidecji_Konduktorzy WHERE Imie = @Imie AND Nazwisko = @Nazwisko AND Nr_Sluzbowy = @Nr_Sluzbowy AND Stanowisko = @Stanowisko;";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, connection, tran))
                    {
                        selectCmd.Parameters.AddWithValue("@Imie", karta.Konduktor.Imie);
                        selectCmd.Parameters.AddWithValue("@Nazwisko", karta.Konduktor.Nazwisko);
                        selectCmd.Parameters.AddWithValue("@Nr_Sluzbowy", karta.Konduktor.Nr_Sluzbowy);
                        selectCmd.Parameters.AddWithValue("@Stanowisko", karta.Konduktor.Stanowisko);
                        object result = selectCmd.ExecuteScalar();
                        Id_Konduktora = Convert.ToInt32(result);

                    }
                    string insertQuery = @"
                                    INSERT INTO Karty_Ewidecji (
                                        Id_Konduktora,
                                        Miesiac,
                                        Rok,
                                        Nominal_M_CA,
                                        Total_Godziny_Lacznie_Z_Odpoczynkiem,
                                        Total_Godziny_Pracy,
                                        Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem,
                                        Total_Liczba_Godzin_Relacji_Pracy,
                                        Total_Liczba_Godzin_Nocnych,
                                        Total_Liczba_NadGodzin_Ogolem_50,
                                        Total_Liczba_NadGodzin_Ogolem_100,
                                        Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50,
                                        Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100,
                                        Suma_Fodzin_Przepracowanych_Plus_Absencja,
                                        Nadgodziny_50,
                                        Nadgodziny_100,
                                        Total_Liczba_Godzin_Absencji,
                                        Total_Sprzedaz_Bilety_Zagranica,
                                        Total_Sprzedaz_Bilety_Kraj,
                                        Total_Sprzedaz_Bilety_Globalne,
                                        Total_Sprzedaz_Bilety_Wartosc_Towarow,
                                        Total_Liczba_Napojow_Awaryjnych,
                                        Ostatnia_Modyfikacja_Data,
                                        Ostatnia_Modyfikacja_Osoba
                                    )
                                    OUTPUT INSERTED.Id_Karty
                                    VALUES (
                                        @Id_Konduktora,
                                        @Miesiac,
                                        @Rok,
                                        @Nominal_M_CA,
                                        @Total_Godziny_Lacznie_Z_Odpoczynkiem,
                                        @Total_Godziny_Pracy,
                                        @Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem,
                                        @Total_Liczba_Godzin_Relacji_Pracy,
                                        @Total_Liczba_Godzin_Nocnych,
                                        @Total_Liczba_NadGodzin_Ogolem_50,
                                        @Total_Liczba_NadGodzin_Ogolem_100,
                                        @Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50,
                                        @Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100,
                                        @Suma_Fodzin_Przepracowanych_Plus_Absencja,
                                        @Nadgodziny_50,
                                        @Nadgodziny_100,
                                        @Total_Liczba_Godzin_Absencji,
                                        @Total_Sprzedaz_Bilety_Zagranica,
                                        @Total_Sprzedaz_Bilety_Kraj,
                                        @Total_Sprzedaz_Bilety_Globalne,
                                        @Total_Sprzedaz_Bilety_Wartosc_Towarow,
                                        @Total_Liczba_Napojow_Awaryjnych,
                                        @Ostatnia_Modyfikacja_Data,
                                        @Ostatnia_Modyfikacja_Osoba
                                    );
                                    ";


                    using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                    {
                        insertCmd.Parameters.AddWithValue("@Id_Konduktora", Id_Konduktora);
                        insertCmd.Parameters.AddWithValue("@Miesiac", karta.Miesiac);
                        insertCmd.Parameters.AddWithValue("@Rok", karta.Rok);
                        insertCmd.Parameters.AddWithValue("@Nominal_M_CA", karta.Nominal_M_CA);
                        insertCmd.Parameters.AddWithValue("@Total_Godziny_Lacznie_Z_Odpoczynkiem", karta.Total_Godziny_Lacznie_Z_Odpoczynkiem);
                        insertCmd.Parameters.AddWithValue("@Total_Godziny_Pracy", karta.Total_Godziny_Pracy);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem", karta.Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Godzin_Relacji_Pracy", karta.Total_Liczba_Godzin_Relacji_Pracy);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Godzin_Nocnych", karta.Total_Liczba_Godzin_Nocnych);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_NadGodzin_Ogolem_50", karta.Total_Liczba_NadGodzin_Ogolem_50);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_NadGodzin_Ogolem_100", karta.Total_Liczba_NadGodzin_Ogolem_100);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50", karta.Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100", karta.Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100);
                        insertCmd.Parameters.AddWithValue("@Suma_Fodzin_Przepracowanych_Plus_Absencja", karta.Suma_Fodzin_Przepracowanych_Plus_Absencja);
                        insertCmd.Parameters.AddWithValue("@Nadgodziny_50", karta.Nadgodziny_50);
                        insertCmd.Parameters.AddWithValue("@Nadgodziny_100", karta.Nadgodziny_100);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Godzin_Absencji", karta.Total_Liczba_Godzin_Absencji);
                        insertCmd.Parameters.AddWithValue("@Total_Sprzedaz_Bilety_Zagranica", karta.Total_Sprzedaz_Bilety_Zagranica);
                        insertCmd.Parameters.AddWithValue("@Total_Sprzedaz_Bilety_Kraj", karta.Total_Sprzedaz_Bilety_Kraj);
                        insertCmd.Parameters.AddWithValue("@Total_Sprzedaz_Bilety_Globalne", karta.Total_Sprzedaz_Bilety_Globalne);
                        insertCmd.Parameters.AddWithValue("@Total_Sprzedaz_Bilety_Wartosc_Towarow", karta.Total_Sprzedaz_Bilety_Wartosc_Towarow);
                        insertCmd.Parameters.AddWithValue("@Total_Liczba_Napojow_Awaryjnych", karta.Total_Liczba_Napojow_Awaryjnych);
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                        int newIdKarty = (int)insertCmd.ExecuteScalar();
                        tran.Commit();
                        foreach (var detale in karta.karta_Ewidencji_Detale)
                        {
                            Insert_Karta_Ewidencji_Detale_To_Db(detale, newIdKarty);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Errors.Add(DateTime.Now.ToString() + " |Konduktor_Reader|Insert to db Error: " + ex.ToString());
                    tran.Rollback();
                }
            }
        }
        private static void Insert_Karta_Ewidencji_Detale_To_Db(Karta_Ewidencji_Detale detale, int Id_Karty)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    string query = @"
                            INSERT INTO Karta_Ewidencji (
                                Id_Karty_Ewidencji,
                                Numer_Relacji,
                                Relacja,
                                Nr_Pociagu,
                                Liczba_Godzin_Relacji_Z_Odpoczynkiem,
                                Liczba_Godzin_Relacji_Pracy,
                                Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50,
                                Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100,
                                Wartosc_Biletow_Zagranica,
                                Wartosc_Biletow_Kraj,
                                Wartosc_Biletow_Globalne,
                                Wartosc_Towarow,
                                Ostatnia_Modyfikacja_Data,
                                Ostatnia_Modyfikacja_Osoba
                            )
                            VALUES (
                                @Id_Karty_Ewidencji,
                                @Numer_Relacji,
                                @Relacja,
                                @Nr_Pociagu,
                                @Liczba_Godzin_Relacji_Z_Odpoczynkiem,
                                @Liczba_Godzin_Relacji_Pracy,
                                @Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50,
                                @Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100,
                                @Wartosc_Biletow_Zagranica,
                                @Wartosc_Biletow_Kraj,
                                @Wartosc_Biletow_Globalne,
                                @Wartosc_Towarow,
                                @Ostatnia_Modyfikacja_Data,
                                @Ostatnia_Modyfikacja_Osoba
                            )";

                    using (SqlCommand command = new SqlCommand(query, connection, tran))
                    {
                        command.Parameters.AddWithValue("@Id_Karty_Ewidencji", Id_Karty);
                        command.Parameters.AddWithValue("@Numer_Relacji", detale.Numer_Relacji);
                        command.Parameters.AddWithValue("@Relacja", detale.Relacja);
                        command.Parameters.AddWithValue("@Nr_Pociagu", detale.Nr_Pociagu);
                        command.Parameters.AddWithValue("@Liczba_Godzin_Relacji_Z_Odpoczynkiem", detale.Liczba_Godzin_Relacji_Z_Odpoczynkiem);
                        command.Parameters.AddWithValue("@Liczba_Godzin_Relacji_Pracy", detale.Liczba_Godzin_Relacji_Pracy);
                        command.Parameters.AddWithValue("@Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50", detale.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50);
                        command.Parameters.AddWithValue("@Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100", detale.Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100);
                        command.Parameters.AddWithValue("@Wartosc_Biletow_Zagranica", detale.Wartosc_Biletow_Zagranica);
                        command.Parameters.AddWithValue("@Wartosc_Biletow_Kraj", detale.Wartosc_Biletow_Kraj);
                        command.Parameters.AddWithValue("@Wartosc_Biletow_Globalne", detale.Wartosc_Biletow_Globalne);
                        command.Parameters.AddWithValue("@Wartosc_Towarow", detale.Wartosc_Towarow);
                        command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                        command.ExecuteNonQuery();
                    }

                    tran.Commit();
                    foreach (var dni in detale.Lista_Detali_Dni)
                    {
                        if (dni.Dzien != 0)
                        {
                            int tmp = Insert_Karta_Ewidecji_Dni_To_Db(Id_Karty, dni);
                            Insert_Godziny_Dnia_To_Db(tmp, dni.Dzien, dni);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Errors.Add(DateTime.Now.ToString() + " |Konduktor_Reader|Insert to db Error: " + ex.ToString());
                    tran.Rollback();
                }
            }
        }
        private static int Insert_Karta_Ewidecji_Dni_To_Db(int id_Relacji, Detale_Dzien dni)
        {
            int insertedId = 0;

            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    string query = @"
                                INSERT INTO Karta_Ewidecji_Dni (
                                    Id_Karta_Ewidencji,
                                    Dzien,
                                    Godziny_Lacznie_Z_Odpoczynkiem,
                                    Godziny_Pracy,
                                    Liczba_Godzin_Nocnych,
                                    Liczba_NadGodzin_Ogolem_50,
                                    Liczba_NadGodzin_Ogolem_100,
                                    Nazwa_Absencji,
                                    Liczba_Godzin_Absencji,
                                    Liczba_Napojow_Awaryjnych,
                                    Ostatnia_Modyfikacja_Data,
                                    Ostatnia_Modyfikacja_Osoba
                                )
                                VALUES (
                                    @Id_Karta_Ewidencji,
                                    @Dzien,
                                    @Godziny_Lacznie_Z_Odpoczynkiem,
                                    @Godziny_Pracy,
                                    @Liczba_Godzin_Nocnych,
                                    @Liczba_NadGodzin_Ogolem_50,
                                    @Liczba_NadGodzin_Ogolem_100,
                                    @Nazwa_Absencji,
                                    @Liczba_Godzin_Absencji,
                                    @Liczba_Napojow_Awaryjnych,
                                    @Ostatnia_Modyfikacja_Data,
                                    @Ostatnia_Modyfikacja_Osoba
                                );
                                SELECT SCOPE_IDENTITY();
                                ";
                    using (SqlCommand command = new SqlCommand(query, connection, tran))
                    {
                        command.Parameters.AddWithValue("@Id_Karta_Ewidencji", id_Relacji);
                        command.Parameters.AddWithValue("@Dzien", dni.Dzien);
                        command.Parameters.AddWithValue("@Godziny_Lacznie_Z_Odpoczynkiem", dni.Godziny_Lacznie_Z_Odpoczynkiem);
                        command.Parameters.AddWithValue("@Godziny_Pracy", dni.Godziny_Pracy);
                        command.Parameters.AddWithValue("@Liczba_Godzin_Nocnych", dni.Liczba_Godzin_Nocnych);
                        command.Parameters.AddWithValue("@Liczba_NadGodzin_Ogolem_50", dni.Liczba_NadGodzin_Ogolem_50);
                        command.Parameters.AddWithValue("@Liczba_NadGodzin_Ogolem_100", dni.Liczba_NadGodzin_Ogolem_100);
                        command.Parameters.AddWithValue("@Nazwa_Absencji", dni.Nazwa_Absencji);
                        command.Parameters.AddWithValue("@Liczba_Godzin_Absencji", dni.Liczba_Godzin_Absencji);
                        command.Parameters.AddWithValue("@Liczba_Napojow_Awaryjnych", dni.Liczba_Napojow_Awaryjnych);
                        command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                        insertedId = Convert.ToInt32(command.ExecuteScalar());
                    }

                    tran.Commit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Errors.Add(DateTime.Now.ToString() + " |Konduktor_Reader|Insert to db Error: " + ex.ToString());
                    tran.Rollback();
                }
            }

            return insertedId;
        }
        private static void Insert_Godziny_Dnia_To_Db(int Id_Karty, int dziennum, Detale_Dzien dzien)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    for (int i = 0; i < dzien.Godziny_Pracy_Od.Count; i++)
                    {
                        string query = @"
                                INSERT INTO Godziny_Pracy_Dnia (
                                    Id_Karta_Ewidecji_Dni,
                                    Dzien,
                                    Godziny_Pracy_Od,
                                    Godziny_Pracy_Do,
                                    Ostatnia_Modyfikacja_Data,
                                    Ostatnia_Modyfikacja_Osoba
                                )
                                VALUES (
                                    @Id_Karty,
                                    @Dzien,
                                    @Godziny_Pracy_Od,
                                    @Godziny_Pracy_Do,
                                    @Ostatnia_Modyfikacja_Data,
                                    @Ostatnia_Modyfikacja_Osoba
                                );
                                ";

                        using (SqlCommand cmd = new SqlCommand(query, connection, tran))
                        {
                            cmd.Parameters.AddWithValue("@Id_Karty", Id_Karty);
                            cmd.Parameters.AddWithValue("@Dzien", dziennum);
                            cmd.Parameters.AddWithValue("@Godziny_Pracy_Od", dzien.Godziny_Pracy_Od[i]);
                            cmd.Parameters.AddWithValue("@Godziny_Pracy_Do", dzien.Godziny_Pracy_Do[i]);
                            cmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            cmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    for (int i = 0; i < dzien.Godziny_Odpoczynku_Od.Count; i++)
                    {
                        string query = @"
                                INSERT INTO Godziny_Odpoczynku_Dnia (
                                    Id_Karta_Ewidecji_Dni,
                                    Dzien,
                                    Godziny_Odpoczynku_Od,
                                    Godziny_Odpoczynku_Do,
                                    Ostatnia_Modyfikacja_Data,
                                    Ostatnia_Modyfikacja_Osoba
                                )
                                VALUES (
                                    @Id_Karty,
                                    @Dzien,
                                    @Godziny_Odpoczynku_Od,
                                    @Godziny_Odpoczynku_Do,
                                    @Ostatnia_Modyfikacja_Data,
                                    @Ostatnia_Modyfikacja_Osoba
                                );
                                ";

                        using (SqlCommand cmd = new SqlCommand(query, connection, tran))
                        {
                            cmd.Parameters.AddWithValue("@Id_Karty", Id_Karty);
                            cmd.Parameters.AddWithValue("@Dzien", dziennum);
                            cmd.Parameters.AddWithValue("@Godziny_Odpoczynku_Od", dzien.Godziny_Odpoczynku_Od[i]);
                            cmd.Parameters.AddWithValue("@Godziny_Odpoczynku_Do", dzien.Godziny_Odpoczynku_Do[i]);
                            cmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            cmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                            cmd.ExecuteNonQuery();
                        }
                    }

                    tran.Commit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    Errors.Add(DateTime.Now.ToString() + " |Konduktor_Reader|Insert to db Error: " + ex.ToString());
                    tran.Rollback();
                }
            }
        }
    }
}
