using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Security.Cryptography;

namespace All_Readeer
{
    internal static class Grafik_Pracy_Reader
    {
        public class Grafik_Pracy
        {
            public string Nazwa_pliku = "";
            public int Nr_zakladki = 0;
            public int Miesiac { get; set; }
            public int Rok { get; set; }
            public int Nominal_Godzin { get; set; }
            public Oddzial Oddzial { get; set; } = new();
            public List<Grafik_Pracy_Legenda> Lista_Grafik_Pracy_Legenda { get; set; } = new();
            public List<DaneMiesiacaPracownika> Lista_Dane_Miesiaca_Pracownika { get; set; } = new();
            public void ParseAndSetMiesiac(string nazwa)
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
                }
            }
        }
        public class Oddzial
        {
            public string Nazwa { get; set; } = "";
        }
        public class DaneMiesiacaPracownika
        {
            public Pracownik pracownik = new();
            public List<Grafik_Detale_Dnia> dni = new();
            public int Total_Godzin = 0;
        }
        public class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
        }
        public class Grafik_Detale_Dnia
        {
            public int Dzien { get; set; }
            public int Kod { get; set; }
        }
        public class Grafik_Pracy_Legenda
        {
            public string Kolor { get; set; } = "";
            public int Kod { get; set; }
            public string Opis { get; set; } = "";
            public int Ilosc_Godzin { get; set; }
            public void LiczGodziny(string przedzial)
            {
                var godziny = przedzial.Split('-');
                int start = int.Parse(godziny[0]);
                int end = int.Parse(godziny[1]);
                if (end >= start)
                {
                    Ilosc_Godzin = end - start;
                }
                else
                {
                    Ilosc_Godzin = (24 - start) + end;
                }
            }
        }
        private static string FilePath = "";
        private static string Connection_String = "Server=ITEGER-NT;Database=nowaBazaZadanie;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
        private static string Optima_Connection_String = "Server=ITEGER-NT;Database=CDN_Wars_5;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
        private static Grafik_Pracy grafik_pracy = new();
        private static IXLWorksheet worksheet = null!;
        private static DateTime Last_Mod_Time = DateTime.Now;
        private static string Last_Mod_Osoba = "";
        public static void Set_File_Path(string NewfilePath)
        {
            FilePath = NewfilePath;
        }
        public static void Set_Errors_File_Folder(string NewfilePath)
        {
            Program.error_logger.Set_Error_File_Path(NewfilePath);
        }
        public static void Set_Optima_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }
        public static void Set_Db_Tables_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }
        public static void Process_For_Db()
        {
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            List<Grafik_Pracy> Grafiki = ReadXlsx();
            foreach (var Grafik in Grafiki)
            {
                List<Pracownik> listaPracowników = new();
                foreach (var prac in Grafik.Lista_Dane_Miesiaca_Pracownika)
                {
                    listaPracowników.Add(prac.pracownik);
                }
                Insert_Pracownicy_To_Db(listaPracowników);
                int Id_Grafiku = Insert_Grafik_To_Db(Grafik);
                Insert_Legenda_Grafiku_To_Db(Id_Grafiku, Grafik.Lista_Grafik_Pracy_Legenda);
                Insert_Grafik_Pracy_Detale_To_Db(Id_Grafiku, Grafik.Lista_Dane_Miesiaca_Pracownika);


            }
        }
        public static void Process_For_Optima()
        {
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            List<Grafik_Pracy> Grafiki = ReadXlsx();
            foreach (var Grafik in Grafiki)
            {
                Wpierdol_Obecnosci_do_Optimy(Grafik);
            }
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
        private static void Wczytaj_Naglowek()
        {

            var cellValue = worksheet.Cell(1, 1).GetFormattedString().Split(" ");
            grafik_pracy.ParseAndSetMiesiac(cellValue[3]);

            int rok;
            if (int.TryParse(cellValue[5], out rok))
            {
                grafik_pracy.Rok = rok;
            }
            else
            {
                grafik_pracy.Rok = 0;
            }

            int nom;
            if (int.TryParse(cellValue[8], out nom))
            {
                grafik_pracy.Nominal_Godzin = nom;
            }
            else
            {
                grafik_pracy.Nominal_Godzin = 0;
            }
        }
        private static void Wczytaj_Oddzial()
        {
            Oddzial oddzial = new();
            var cellValue = worksheet.Cell(2, 1).GetFormattedString();
            while (cellValue.Contains("."))
            {
                cellValue = cellValue.Replace(".", "");
            }
            int index = cellValue.IndexOf("ODDZIAŁ");
            if (index != -1)
            {
                oddzial.Nazwa = cellValue.Substring(index + "ODDZIAŁ".Length).Trim();
            }
            grafik_pracy.Oddzial = oddzial;
        }
        private static int Wczytaj_Legenda()
        {
            List<Grafik_Pracy_Legenda> listalegenda = new();
            var row = 1;
            while (worksheet.Cell(row, 1).GetFormattedString() != "Legenda:")
            {
                row++;
            }
            row++;
            var StartLegendaRow = row;
            var tmpkod = -1;
            while (true)
            {
                try
                {
                    Grafik_Pracy_Legenda legenda = new();
                    var kolor = worksheet.Cell(row, 1).Style.Fill.BackgroundColor.Color.Name.ToString();
                    legenda.Kod = tmpkod;
                    tmpkod--;
                    legenda.Opis = worksheet.Cell(row, 2).GetFormattedString();
                    legenda.Ilosc_Godzin = 0;
                    listalegenda.Add(legenda);
                    row++;
                }
                catch
                {
                    break;
                }
            }

            var col = 4;
            bool end = false;
            int maxkod = 0;
            while (end == false)
            {
                row = StartLegendaRow;
                for (int i = 0; i < 3; i++)
                {
                    var ntmpkod = worksheet.Cell(row, col).GetFormattedString();
                    var ntempopis = worksheet.Cell(row, col + 1).GetFormattedString();
                    if (string.IsNullOrEmpty(ntmpkod))
                    {
                        end = true;
                        break;
                    }
                    Grafik_Pracy_Legenda legenda = new();
                    legenda.Kod = tmpkod;
                    int tmpval;
                    if (int.TryParse(ntmpkod, out tmpval))
                    {
                        legenda.Kod = tmpval;
                    }
                    maxkod = legenda.Kod;
                    legenda.Opis = ntempopis;
                    row++;
                    legenda.LiczGodziny(ntempopis);
                    listalegenda.Add(legenda);
                }
                col += 3;
            }
            Grafik_Pracy_Legenda ulegenda = new();
            ulegenda.Kod = maxkod + 1;
            ulegenda.Opis = "U(urlop wypoczynkowy)";
            ulegenda.Ilosc_Godzin = 0;
            listalegenda.Add(ulegenda);
            grafik_pracy.Lista_Grafik_Pracy_Legenda = listalegenda;
            return maxkod + 2;
        }
        private static void Wczytaj_Dane_Grafiku(int maxkod)
        {
            int row = 4;
            int col = 1;
            int prevDay = 0;
            List<DaneMiesiacaPracownika> listadaneMiesiacaPracownika = new();
            while (!string.IsNullOrEmpty(worksheet.Cell(row, 1).GetFormattedString()))
            {
                col = 1;
                prevDay = 0;
                DaneMiesiacaPracownika daneMiesiacaPracownika = new();
                List<Grafik_Detale_Dnia> detale_Dnia = new();

                Pracownik pracownik = new();
                var namesurname = worksheet.Cell(row, col).GetFormattedString().Split(" ");
                if (string.IsNullOrEmpty(namesurname[0]))
                {
                    Program.error_logger.New_Error(namesurname[0], "NameSurname", col, row);
                    Console.WriteLine(Program.error_logger.Get_Error_String());
                    continue;
                }
                pracownik.Imie = namesurname[0];
                pracownik.Nazwisko = namesurname[1];
                col += 2;
                int currentSheetDay = 0;
                while (true)
                {
                    int tmpval;
                    if (int.TryParse(worksheet.Cell(3, col).GetFormattedString(), out tmpval))
                    {
                        currentSheetDay = tmpval;
                    }
                    if (prevDay + 1 != currentSheetDay)
                    {
                        break;
                    }

                    Grafik_Detale_Dnia dzien = new();
                    dzien.Dzien = currentSheetDay;
                    var cell = worksheet.Cell(row, col).GetFormattedString();

                    if (!string.IsNullOrEmpty(cell))
                    {
                        if (cell == "U")
                        {
                            dzien.Kod = maxkod + 1;
                        }
                        else
                        {
                            int tmpval3;
                            if (int.TryParse(cell, out tmpval3))
                            {
                                dzien.Kod = tmpval3;
                            }
                        }
                    }
                    else
                    {
                        try
                        {
                            string backgroundColor = worksheet.Cell(row, col).Style.Fill.BackgroundColor.Color.Name.ToString();
                            var legenda = grafik_pracy.Lista_Grafik_Pracy_Legenda
                                .FirstOrDefault(l => l.Kolor == backgroundColor);
                            dzien.Kod = legenda?.Kod ?? 0;
                        }
                        catch
                        {
                            dzien.Kod = 0;
                        }
                    }
                    detale_Dnia.Add(dzien);
                    prevDay++;
                    col++;
                }
                daneMiesiacaPracownika.pracownik = pracownik;
                daneMiesiacaPracownika.dni = detale_Dnia;
                if (!string.IsNullOrEmpty(worksheet.Cell(row, col).GetFormattedString()))
                {
                    var totalgodzin = worksheet.Cell(row, col).GetFormattedString().Split('(');
                    int tmpval2;

                    if (totalgodzin.Length > 0 && int.TryParse(totalgodzin[0], out tmpval2))
                    {
                        daneMiesiacaPracownika.Total_Godzin = tmpval2;
                    }
                    else
                    {
                        daneMiesiacaPracownika.Total_Godzin = Oblicz_Total_Godziny(daneMiesiacaPracownika);
                    }
                }
                else
                {
                    daneMiesiacaPracownika.Total_Godzin = Oblicz_Total_Godziny(daneMiesiacaPracownika);
                }
                listadaneMiesiacaPracownika.Add(daneMiesiacaPracownika);
                row++;
            }
            grafik_pracy.Lista_Dane_Miesiaca_Pracownika = listadaneMiesiacaPracownika;
        }
        private static int Oblicz_Total_Godziny(DaneMiesiacaPracownika daneMiesiacaPracownika)
        {
            int suma = 0;
            foreach (var dzien in daneMiesiacaPracownika.dni)
            {
                var legenda = grafik_pracy.Lista_Grafik_Pracy_Legenda.FirstOrDefault(leg => leg.Kod == dzien.Kod);
                if (legenda != null)
                {
                    suma += legenda.Ilosc_Godzin;
                }
            }
            return suma;
        }
        private static List<Grafik_Pracy> ReadXlsx()
        {
            Program.error_logger.Nazwa_Pliku = FilePath;
            List<Grafik_Pracy> grafiki_Pracy = new();
            using (var workbook = new XLWorkbook(FilePath))
            {
                for (var i = 1; i <= workbook.Worksheets.Count; i++)
                {
                    Program.error_logger.Nr_Zakladki = i;
                    worksheet = workbook.Worksheet(i);
                    grafik_pracy = new();
                    grafik_pracy.Nazwa_pliku = FilePath;
                    grafik_pracy.Nr_zakladki = i;
                    Wczytaj_Naglowek();
                    Wczytaj_Oddzial();
                    int ilosclegend = Wczytaj_Legenda();
                    Wczytaj_Dane_Grafiku(ilosclegend);
                    grafiki_Pracy.Add(grafik_pracy);
                }
            }
            return grafiki_Pracy;
        }
        private static void Insert_Pracownicy_To_Db(List<Pracownik> pracownicy)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();

                try
                {
                    foreach (var pracownik in pracownicy)
                    {
                        string checkQuery = "SELECT COUNT(1) FROM Grafik_Pracy_Pracownicy WHERE Imie = @Imie AND Nazwisko = @Nazwisko";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection, tran))
                        {
                            checkCmd.Parameters.AddWithValue("@Imie", pracownik.Imie);
                            checkCmd.Parameters.AddWithValue("@Nazwisko", pracownik.Nazwisko);

                            int count = (int)checkCmd.ExecuteScalar();

                            if (count == 0)
                            {
                                string insertQuery = "INSERT INTO Grafik_Pracy_Pracownicy (Imie, Nazwisko, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) VALUES (@Imie, @Nazwisko, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba)";
                                using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                                {
                                    insertCmd.Parameters.AddWithValue("@Imie", pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@Nazwisko", pracownik.Nazwisko);
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
                    tran.Rollback();
                }
            }
        }
        private static int Insert_Oddzial_To_Db(Oddzial oddzial)
        {
            int Id_Oddzialu = 0;
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    try
                    {
                        string checkQuery = "SELECT Id_Oddzialu FROM Grafik_Pracy_Oddzialy WHERE Nazwa = @Nazwa";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection, tran))
                        {
                            checkCmd.Parameters.Add("@Nazwa", SqlDbType.NVarChar).Value = oddzial.Nazwa;

                            var result = checkCmd.ExecuteScalar();

                            if (result != null)
                            {
                                Id_Oddzialu = Convert.ToInt32(result);
                            }
                            else
                            {
                                string insertQuery = "INSERT INTO Grafik_Pracy_Oddzialy (Nazwa, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) VALUES (@Nazwa, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba); SELECT SCOPE_IDENTITY();";
                                using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                                {
                                    insertCmd.Parameters.Add("@Nazwa", SqlDbType.NVarChar).Value = oddzial.Nazwa;
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                                    Id_Oddzialu = Convert.ToInt32(insertCmd.ExecuteScalar());
                                }
                            }
                        }
                        tran.Commit();
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            return Id_Oddzialu;
        }
        private static int Insert_Grafik_To_Db(Grafik_Pracy grafik_pracy)
        {
            int insertedId = -1;
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    int Id_Oddzialu = Insert_Oddzial_To_Db(grafik_pracy.Oddzial);
                    string insertQuery = "INSERT INTO Grafiki_Pracy (Miesiac, Rok, Nominal_Godzin, Id_Oddzialu, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                            "VALUES (@Miesiac, @Rok, @Nominal_Godzin, @Id_Oddzialu, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba); " +
                                            "SELECT SCOPE_IDENTITY();";
                    using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                    {
                        insertCmd.Parameters.Add("@Miesiac", SqlDbType.Int).Value = grafik_pracy.Miesiac;
                        insertCmd.Parameters.Add("@Rok", SqlDbType.Int).Value = grafik_pracy.Rok;
                        insertCmd.Parameters.Add("@Nominal_Godzin", SqlDbType.Int).Value = grafik_pracy.Nominal_Godzin;
                        insertCmd.Parameters.Add("@Id_Oddzialu", SqlDbType.Int).Value = Id_Oddzialu;
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                        insertedId = Convert.ToInt32(insertCmd.ExecuteScalar());
                        tran.Commit();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    tran.Rollback();
                }
            }
            return insertedId;
        }
        private static void Insert_Legenda_Grafiku_To_Db(int ID_Grafiku_Pracy, List<Grafik_Pracy_Legenda> legendy)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    string insertQuery = "INSERT INTO Grafik_Pracy_Legenda (Id_Grafiku, Id_Kodu, Opis, Ilosc_Godzin, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                            "VALUES (@Id_Grafiku, @Id_Kodu, @Opis, @Ilosc_Godzin, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba);";
                    foreach (var legenda in legendy)
                    {
                        using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@Id_Grafiku", ID_Grafiku_Pracy);
                            insertCmd.Parameters.AddWithValue("@Id_Kodu", legenda.Kod);
                            insertCmd.Parameters.AddWithValue("@Opis", legenda.Opis);
                            insertCmd.Parameters.AddWithValue("@Ilosc_Godzin", legenda.Ilosc_Godzin);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba); insertCmd.ExecuteNonQuery();
                        }
                    }
                    tran.Commit();
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    tran.Rollback();
                }
            }
        }
        private static void Insert_Grafik_Pracy_Detale_To_Db(int ID_Grafiku_Pracy, List<DaneMiesiacaPracownika> daneMiesiaca)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    try
                    {
                        foreach (var dana in daneMiesiaca)
                        {
                            string checkQuery = "SELECT Id_Pracownika FROM Grafik_Pracy_Pracownicy WHERE Imie = @Imie AND Nazwisko = @Nazwisko";
                            int Id_Pracownika = 0;

                            using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection, tran))
                            {
                                checkCmd.Parameters.AddWithValue("@Imie", dana.pracownik.Imie);
                                checkCmd.Parameters.AddWithValue("@Nazwisko", dana.pracownik.Nazwisko);

                                object result = checkCmd.ExecuteScalar();
                                if (result != null)
                                {
                                    Id_Pracownika = Convert.ToInt32(result);
                                }
                            }
                            string insertQuery = "INSERT INTO Grafik_Pracy_Detale (Id_Grafiku, Id_Pracownika, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) VALUES (@Id_Grafiku, @Id_Pracownika, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba); " +
                                                    "SELECT SCOPE_IDENTITY();";
                            using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@Id_Grafiku", ID_Grafiku_Pracy);
                                insertCmd.Parameters.AddWithValue("@Id_Pracownika", Id_Pracownika);
                                insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                                insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                                int Id_Detalu = Convert.ToInt32(insertCmd.ExecuteScalar());

                                foreach (var dzien in dana.dni)
                                {
                                    Insert_Grafik_Pracy_Detale_Dni_To_Db(Id_Detalu, dzien, connection, tran);
                                }
                            }
                        }
                        tran.Commit();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        tran.Rollback();
                    }
                }
            }
        }
        private static void Insert_Grafik_Pracy_Detale_Dni_To_Db(int Id_Detalu, Grafik_Detale_Dnia detale_Dnia, SqlConnection connection, SqlTransaction tran)
        {
            string insertQuery = "INSERT INTO Grafik_Pracy_Detale_Dni (Id_Detalu, Dzien, Id_Kodu, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                    "VALUES (@Id_Detalu, @Dzien, @Id_Kodu, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba);";
            using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
            {
                insertCmd.Parameters.AddWithValue("@Id_Detalu", Id_Detalu);
                insertCmd.Parameters.AddWithValue("@Dzien", detale_Dnia.Dzien);
                insertCmd.Parameters.AddWithValue("@Id_Kodu", detale_Dnia.Kod);
                insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                insertCmd.ExecuteNonQuery();
            }
        }
        private static void Wpierdol_Obecnosci_do_Optimy(Grafik_Pracy grafik)
        {
            using (SqlConnection connection = new SqlConnection(Optima_Connection_String))
            {
                SqlTransaction tran = connection.BeginTransaction();
                foreach (var garfikD in grafik.Lista_Dane_Miesiaca_Pracownika)
                {
                    foreach (var dzien in garfikD.dni)
                    {

                        var sqlQuery = $@"
            DECLARE @id int;

            -- dodawaina pracownika do pracx i init pracpracdni
            DECLARE @PRI_PraId INT = (SELECT DISTINCT PRI_PraId FROM CDN.Pracidx where PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Imie1 = @PracownikImieInsert and PRI_Typ = 1)

            IF @PRI_PraId IS NULL
            BEGIN
                THROW 50000, 'Brak takiego pracownika w bazie' + @PracownikNazwiskoInsert + ' ' + @PracownikImieInsert, 1;
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

                        // brak obecości
                        if (dzien.Kod == 0)
                        {
                            continue;
                        }

                        var godziny = "0-0";
                        var opis = grafik.Lista_Grafik_Pracy_Legenda
                            .Where(kod => kod.Kod == dzien.Kod)
                            .Select(kod => kod.Opis)
                            .FirstOrDefault();

                        if (!string.IsNullOrEmpty(opis))
                        {
                            godziny = opis;
                            var TeoretyczneGodzPracy = Math.Abs(int.Parse(godziny.Split('-')[0]) - int.Parse(godziny.Split('-')[1]));
                            connection.Open();
                            try
                            {
                                using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
                                {
                                    insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{grafik.Rok}-{grafik.Miesiac:D2}-{dzien.Dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = "1899-12-30" + ' ' + godziny.Split('-')[0] + ":00";
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = "1899-12-30" + ' ' + godziny.Split('-')[1] + ":00";
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", TeoretyczneGodzPracy);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", TeoretyczneGodzPracy);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", garfikD.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", garfikD.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                                    insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                                    insertCmd.ExecuteScalar();
                                }
                            }
                            catch (SqlException ex)
                            {
                                Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                                Console.WriteLine(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                                tran.Rollback();
                                var e = new Exception(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                                e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                                throw e;
                            }
                        }
                    }
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Poprawnie dodawno plan z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                Console.ForegroundColor = ConsoleColor.White;
                tran.Commit();
            }
        }
    }
}