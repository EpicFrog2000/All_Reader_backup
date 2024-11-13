using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Globalization;

namespace All_Readeer
{
    internal class Oddzial_Reader
    {
        private class Stawki_Wynagrodzeń_Rzyczałtowych_Detale
        {
            public string Nr_Relacji_Detal { get; set; } = "";
            public string Opis { get; set; } = "";
            public decimal Czas_Relacji_Calkowity { get; set; }
            public decimal Czas_Pracy_Ogolem { get; set; }
            public decimal Czas_Pracy_Podstawowe { get; set; }
            public decimal Godziny_Nadliczbowe_50 { get; set; }
            public decimal Godziny_Nadliczbowe_100 { get; set; }
            public decimal Godziny_Pracy_W_Nocy { get; set; }
            public decimal Czas_Odpoczynku { get; set; }
            public decimal Podstawowa_Stawka_Godzinowa { get; set; }
            public decimal Podstawowe_Wynagrodzenie_Ryczaltowe { get; set; }
            public decimal Wynagrodzenie_Za_Godziny_NadLiczbowe { get; set; }
            public decimal Dodatek_Za_Prace_W_Nocy { get; set; }
            public decimal Wynagrodzenie_Ryczaltowe_Calkowite { get; set; }
            public decimal Dodatek_Wyjazdowy { get; set; }
        }
        private class Stawki_Wynagrodzeń_Rzyczałtowych_Relacje
        {
            public string Nr_Relacji { get; set; } = "";
            public string Opis_Relacji1 { get; set; } = "";
            public string Opis_Relacji2 { get; set; } = "";
            public string Rocznik { get; set; } = "";
            public List<Stawki_Wynagrodzeń_Rzyczałtowych_Detale> Stawki { get; set; } = [];
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

        private static List<Stawki_Wynagrodzeń_Rzyczałtowych_Relacje> listastawek = [];
        private static List<string> Errors = new();
        private static string FilePath = "";
        private static ErrorLogger Error_Logger = new(); // TODO: <- jest tylko zadeklarowany, dopisz wykrywanie błędów w trakcie czytania pliku itp
        private static string Connection_String = "Server=ITEGER-NT;Database=nowaBazaZadanie;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
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
            try
            {
                ReadXlsx();
                foreach (var item in listastawek)
                {
                    Insert_Data_To_Db(item);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Errors.Add(DateTime.Now.ToString() + " |Oddzial_Reader|Program Error: " + ex.ToString());
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

        private static void ReadXlsx()
        {
            using (var workbook = new XLWorkbook(FilePath))
            {
                foreach (var worksheet in workbook.Worksheets)
                {
                    try
                    {
                        int row = 8;
                        var cellValue = worksheet.Cell(1, 3)?.GetFormattedString();
                        var Rocznik = "";

                        if (!string.IsNullOrEmpty(cellValue))
                        {
                            Rocznik = cellValue.Split(' ')[^1];
                        }
                        Stawki_Wynagrodzeń_Rzyczałtowych_Relacje DANE = new();
                        DANE.Rocznik = Rocznik;
                        while (!worksheet.Cell(row, 1).IsEmpty())
                        {
                            if (worksheet.Cell(row, 1).Style.Font.Bold)
                            {
                                if (row != 8)
                                {
                                    listastawek.Add(DANE);
                                }
                                DANE = new();
                                DANE.Rocznik = Rocznik;
                                DANE.Nr_Relacji = worksheet.Cell(row, 1).GetString();
                                DANE.Opis_Relacji1 = worksheet.Cell(row, 2).GetFormattedString();
                                row++;
                                DANE.Opis_Relacji2 = worksheet.Cell(row, 2).GetFormattedString();
                            }
                            else
                            {
                                Stawki_Wynagrodzeń_Rzyczałtowych_Detale detale = new();
                                detale.Nr_Relacji_Detal = worksheet.Cell(row, 1).GetFormattedString();
                                detale.Opis = worksheet.Cell(row, 2).GetFormattedString();
                                detale.Czas_Relacji_Calkowity = Convert.ToDecimal(worksheet.Cell(row, 3).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Czas_Pracy_Ogolem = Convert.ToDecimal(worksheet.Cell(row, 4).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Czas_Pracy_Podstawowe = Convert.ToDecimal(worksheet.Cell(row, 5).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Godziny_Nadliczbowe_50 = Convert.ToDecimal(worksheet.Cell(row, 6).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Godziny_Nadliczbowe_100 = Convert.ToDecimal(worksheet.Cell(row, 7).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Godziny_Pracy_W_Nocy = Convert.ToDecimal(worksheet.Cell(row, 8).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Czas_Odpoczynku = Convert.ToDecimal(worksheet.Cell(row, 9).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Podstawowa_Stawka_Godzinowa = Convert.ToDecimal(worksheet.Cell(row, 10).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Podstawowe_Wynagrodzenie_Ryczaltowe = Convert.ToDecimal(worksheet.Cell(row, 11).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Wynagrodzenie_Za_Godziny_NadLiczbowe = Convert.ToDecimal(worksheet.Cell(row, 12).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Dodatek_Za_Prace_W_Nocy = Convert.ToDecimal(worksheet.Cell(row, 13).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Wynagrodzenie_Ryczaltowe_Calkowite = Convert.ToDecimal(worksheet.Cell(row, 14).GetFormattedString(), CultureInfo.InvariantCulture);
                                detale.Dodatek_Wyjazdowy = Convert.ToDecimal(worksheet.Cell(row, 15).GetFormattedString(), CultureInfo.InvariantCulture);
                                DANE.Stawki.Add(detale);
                            }
                            row++;
                        }
                        listastawek.Add(DANE);
                    }
                    catch(Exception ex)
                    {
                        Errors.Add(DateTime.Now.ToString() + " |Oddzial_Reader|Reading Error: " + ex.ToString());
                    }
                }
            }
        }
        private static void Insert_Data_To_Db(Stawki_Wynagrodzeń_Rzyczałtowych_Relacje relacja)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction transaction = connection.BeginTransaction();
                try
                {
                    int AutoIncremevtValue = 0;
                    string insertRelacjaQuery = @"
                INSERT INTO Stawki_Wynagrodzen_Ryczaltowych_Relacje
                (Nr_Relacji, Opis_Relacji1, Opis_Relacji2, Rocznik, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba)
                VALUES (@Nr_Relacji, @Opis_Relacji1, @Opis_Relacji2, @Rocznik, @Ostatnia_Modyfikacja_Data, @Ostatnia_Modyfikacja_Osoba);
                SELECT SCOPE_IDENTITY();";

                    using (SqlCommand command = new SqlCommand(insertRelacjaQuery, connection, transaction))
                    {
                        command.Parameters.AddWithValue("@Nr_Relacji", relacja.Nr_Relacji);
                        command.Parameters.AddWithValue("@Opis_Relacji1", relacja.Opis_Relacji1);
                        command.Parameters.AddWithValue("@Opis_Relacji2", relacja.Opis_Relacji2);
                        command.Parameters.AddWithValue("@Rocznik", relacja.Rocznik);
                        command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                        AutoIncremevtValue = Convert.ToInt32(command.ExecuteScalar());
                    }

                    string insertDetalQuery = @$"
                INSERT INTO Stawki_Wynagrodzen_Ryczaltowych_Detale
                (Nr_Relacji_Detal, Opis, Czas_Relacji_Calkowity, Czas_Pracy_Ogolem, Czas_Pracy_Podstawowe,
                Godziny_Nadliczbowe_50, Godziny_Nadliczbowe_100, Godziny_Pracy_W_Nocy, Czas_Odpoczynku,
                Podstawowa_Stawka_Godzinowa, Podstawowe_Wynagrodzenie_Ryczaltowe, Wynagrodzenie_Za_Godziny_NadLiczbowe,
                Dodatek_Za_Prace_W_Nocy, Wynagrodzenie_Ryczaltowe_Calkowite, Dodatek_Wyjazdowy, Id_Relacji, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba)
                VALUES
                (@Nr_Relacji_Detal, @Opis, @Czas_Relacji_Calkowity, @Czas_Pracy_Ogolem, @Czas_Pracy_Podstawowe,
                @Godziny_Nadliczbowe_50, @Godziny_Nadliczbowe_100, @Godziny_Pracy_W_Nocy, @Czas_Odpoczynku,
                @Podstawowa_Stawka_Godzinowa, @Podstawowe_Wynagrodzenie_Ryczaltowe, @Wynagrodzenie_Za_Godziny_NadLiczbowe,
                @Dodatek_Za_Prace_W_Nocy, @Wynagrodzenie_Ryczaltowe_Calkowite, @Dodatek_Wyjazdowy, @Id_Relacji, @Ostatnia_Modyfikacja_Data, @Ostatnia_Modyfikacja_Osoba)";

                    foreach (var detal in relacja.Stawki)
                    {
                        using (SqlCommand command = new SqlCommand(insertDetalQuery, connection, transaction))
                        {
                            command.Parameters.AddWithValue("@Nr_Relacji_Detal", detal.Nr_Relacji_Detal);
                            command.Parameters.AddWithValue("@Opis", detal.Opis);
                            command.Parameters.AddWithValue("@Czas_Relacji_Calkowity", detal.Czas_Relacji_Calkowity);
                            command.Parameters.AddWithValue("@Czas_Pracy_Ogolem", detal.Czas_Pracy_Ogolem);
                            command.Parameters.AddWithValue("@Czas_Pracy_Podstawowe", detal.Czas_Pracy_Podstawowe);
                            command.Parameters.AddWithValue("@Godziny_Nadliczbowe_50", detal.Godziny_Nadliczbowe_50);
                            command.Parameters.AddWithValue("@Godziny_Nadliczbowe_100", detal.Godziny_Nadliczbowe_100);
                            command.Parameters.AddWithValue("@Godziny_Pracy_W_Nocy", detal.Godziny_Pracy_W_Nocy);
                            command.Parameters.AddWithValue("@Czas_Odpoczynku", detal.Czas_Odpoczynku);
                            command.Parameters.AddWithValue("@Podstawowa_Stawka_Godzinowa", detal.Podstawowa_Stawka_Godzinowa);
                            command.Parameters.AddWithValue("@Podstawowe_Wynagrodzenie_Ryczaltowe", detal.Podstawowe_Wynagrodzenie_Ryczaltowe);
                            command.Parameters.AddWithValue("@Wynagrodzenie_Za_Godziny_NadLiczbowe", detal.Wynagrodzenie_Za_Godziny_NadLiczbowe);
                            command.Parameters.AddWithValue("@Dodatek_Za_Prace_W_Nocy", detal.Dodatek_Za_Prace_W_Nocy);
                            command.Parameters.AddWithValue("@Wynagrodzenie_Ryczaltowe_Calkowite", detal.Wynagrodzenie_Ryczaltowe_Calkowite);
                            command.Parameters.AddWithValue("@Dodatek_Wyjazdowy", detal.Dodatek_Wyjazdowy);
                            command.Parameters.AddWithValue("@Id_Relacji", AutoIncremevtValue);
                            command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            command.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                            command.ExecuteNonQuery();
                        }
                    }
                    transaction.Commit();
                }
                catch (Exception e)
                {
                    transaction.Rollback();
                    Errors.Add(DateTime.Now.ToString() + " |Oddzial_Reader|Insert to db Error: " + e.ToString());
                }
            }
        }
    }
}
