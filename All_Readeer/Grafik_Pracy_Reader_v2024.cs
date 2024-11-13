using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
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
        private string File_Path = "";
        private string Last_Mod_Osoba = "";
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Connection_String = "";
        private string Optima_Connection_String = "";
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time)
        {
            Grafik grafik = new();
            grafik.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
            grafik.nr_zakladki = Program.error_logger.Nr_Zakladki;
            Get_Header_Karta_Info(worksheet, ref grafik);
            Get_Dane_Dni(worksheet, ref grafik);
            Wpierdol_Plan_do_Optimy(grafik);
        }
        public void Set_File_Path(string New_File_Path)
        {
            if (string.IsNullOrEmpty(New_File_Path))
            {
                Console.WriteLine("error: Empty file path");
                return;
            }
            File_Path = New_File_Path;
        }
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
        public void Set_Db_Tables_ConnectionString(string NewConnectionString)
        {
            if (string.IsNullOrEmpty(NewConnectionString))
            {
                Console.WriteLine("error: Empty Connection string");
                return;
            }
            Connection_String = NewConnectionString;
        }
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            if (string.IsNullOrEmpty(NewConnectionString))
            {
                Console.WriteLine("error: Empty Connection string");
                return;
            }
            Optima_Connection_String = NewConnectionString;
        }
        public void Process()
        {
            List<Grafik> grafiki = ReadXlsx();
            List<Pracownik> listaPracowników = new();
            foreach (var grafik in grafiki)
            {
                foreach (var danedni in grafik.dane_dni)
                {
                    listaPracowników.Add(danedni.pracownik);
                }
            }
            try
            {
                Insert_Pracownicy_To_Db(listaPracowników);
                foreach (var grafik in grafiki)
                {
                    int Id_Grafiku = Insert_Grafik_To_Db(grafik);
                    Insert_Grafik_Pracy_Detale_To_Db(Id_Grafiku, grafik.dane_dni);
                    Wpierdol_Plan_do_Optimy(grafik);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
        private List<Grafik> ReadXlsx()
        {
            Program.error_logger.Nazwa_Pliku = File_Path;
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            List<Grafik> grafiki = [];
            try
            {
                using (var workbook = new XLWorkbook(File_Path))
                {
                    Program.error_logger.Nr_Zakladki = workbook.Worksheets.Count;
                    var worksheet = workbook.Worksheet(workbook.Worksheets.Count);
                    try
                    {
                        Grafik grafik = new Grafik();
                        grafik.nazwa_pliku = File_Path;
                        grafik.nr_zakladki = workbook.Worksheets.Count;
                        Get_Header_Karta_Info(worksheet, ref grafik);
                        Get_Dane_Dni(worksheet, ref grafik);
                        grafiki.Add(grafik);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return grafiki;
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
                if(string.IsNullOrEmpty(nazwiskoimie))
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
                    Program.error_logger.New_Error(nazwiskoimie, "Nazwisko i Imie", poz.col, poz.row, ex.Message);
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
                        dane_dnia.dzien = int.Parse(dziennr);
                        // get godziny pracy dnia
                        for (int k = 1; k <= height; k++)
                        {
                            Godz_Pracy godziny = new();
                            var godzr = worksheet.Cell(3 + k, poz.col + 1 + j).GetValue<string>().Trim();
                            if (!string.IsNullOrEmpty(godzr) && godzr != "" && godzr.Length > 0)
                            {
                                if (DateTime.TryParse(godzr, out DateTime parsedTime))
                                {
                                    godziny.Godz_Rozpoczecia_Pracy = parsedTime.TimeOfDay;
                                }
                                else
                                {
                                    Program.error_logger.New_Error(godzr, "godziny rozpoczęcia pracy.", poz.col, poz.row, "Nieprawidłowy format godziny rozpoczęcia pracy.");
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }
                            }
                            var godzz = worksheet.Cell(3 + k, poz.col + 1 + j + 1).GetValue<string>().Trim();
                            if (!string.IsNullOrEmpty(godzz) && godzz != "" && godzz.Length > 0)
                            {
                                if (DateTime.TryParse(godzz, out DateTime parsedTime))
                                {
                                    godziny.Godz_Zakonczenia_Pracy = parsedTime.TimeOfDay;
                                }
                                else
                                {
                                    Program.error_logger.New_Error(godzr, "godziny zakończenia pracy.", poz.col, poz.row, "Nieprawidłowy format godziny zakończenia pracy.");
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }
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
        private int Insert_Grafik_To_Db(Grafik grafik_pracy)
        {
            int insertedId = -1;
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    string insertQuery = "INSERT INTO Grafiki_Pracy (Miesiac, Rok, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                            "VALUES (@Miesiac, @Rok, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba); " +
                                            "SELECT SCOPE_IDENTITY();";
                    using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                    {
                        insertCmd.Parameters.Add("@Miesiac", SqlDbType.Int).Value = grafik_pracy.miesiac;
                        insertCmd.Parameters.Add("@Rok", SqlDbType.Int).Value = grafik_pracy.rok;
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                        insertedId = Convert.ToInt32(insertCmd.ExecuteScalar());
                        tran.Commit();
                    }
                }
                catch (Exception ex)
                {
                    Program.error_logger.New_Custom_Error(ex.Message);
                    tran.Rollback();
                }
            }
            return insertedId;
        }
        private void Insert_Grafik_Pracy_Detale_To_Db(int ID_Grafiku_Pracy, List<Dane_Dni> daneMiesiaca)
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
                            string insertQuery = "INSERT INTO Grafik_Pracy_Detale (Id_Grafiku, Id_Pracownika, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                                 "VALUES (@Id_Grafiku, @Id_Pracownika, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba); " +
                                                 "SELECT SCOPE_IDENTITY();";
                            using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@Id_Grafiku", ID_Grafiku_Pracy);
                                insertCmd.Parameters.AddWithValue("@Id_Pracownika", Id_Pracownika);
                                insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                                insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                                int Id_Detalu = Convert.ToInt32(insertCmd.ExecuteScalar());

                                foreach (var dzien in dana.dane_dnia)
                                {
                                    Insert_Grafik_Pracy_Detale_Dni_To_Db(Id_Detalu, dzien, connection, tran);
                                }
                            }
                        }
                        tran.Commit();
                    }
                    catch (Exception ex)
                    {
                        Program.error_logger.New_Custom_Error(ex.Message + " Nie wpisano dnia do bazy");
                        Console.WriteLine(Program.error_logger.Get_Error_String());
                        tran.Rollback();
                    }
                }
            }

        }
        private void Insert_Grafik_Pracy_Detale_Dni_To_Db(int Id_Detalu, Dane_Dnia detale_Dnia, SqlConnection connection, SqlTransaction tran)
        {
            string insertQuery = "INSERT INTO Grafik_Pracy_Detale_Dni (Id_Detalu, Dzien, Godzina_Rozpoczecia_Pracy, Godzina_Zakonczenia_Pracy, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                    "VALUES (@Id_Detalu, @Dzien, @Godzina_Rozpoczecia_Pracy, @Godzina_Zakonczenia_Pracy, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Osoba);";
            foreach (var godziny in detale_Dnia.godz_pracy)
            {
                using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                {
                    insertCmd.Parameters.AddWithValue("@Id_Detalu", Id_Detalu);
                    insertCmd.Parameters.AddWithValue("@Dzien", detale_Dnia.dzien);
                    insertCmd.Parameters.AddWithValue("@Godzina_Rozpoczecia_Pracy", godziny.Godz_Rozpoczecia_Pracy);
                    insertCmd.Parameters.AddWithValue("@Godzina_Zakonczenia_Pracy", godziny.Godz_Zakonczenia_Pracy);
                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Osoba", Last_Mod_Osoba);
                    insertCmd.ExecuteNonQuery();
                }
            }
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
                    Program.error_logger.New_Custom_Error(ex.Message);
                    tran.Rollback();
                }
            }
        }
        private void Wpierdol_Plan_do_Optimy(Grafik grafik)
        {
            var sqlQuery = $@"
DECLARE @id int;

DECLARE @PRI_PraId INT = (SELECT DISTINCT PRI_PraId FROM CDN.Pracidx where PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Imie1 = @PracownikImieInsert and PRI_Typ = 1)

IF @PRI_PraId IS NULL
BEGIN
DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie: ' + @PracownikNazwiskoInsert + ' ' + @PracownikImieInsert;
THROW 50000, @ErrorMessage, 1;
END

DECLARE @EXISTSPRACTEST INT = (SELECT CDN.PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId)

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

DECLARE @PRA_PraId INT = (SELECT cdn.PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId);

DECLARE @EXISTSDZIEN INT = (SELECT COUNT([CDN].[PracPlanDni].[PPL_Data]) FROM cdn.PracPlanDni WHERE cdn.PracPlanDni.PPL_PraId = @PRA_PraId and [CDN].[PracPlanDni].[PPL_Data] = @DataInsert)
IF @EXISTSDZIEN = 0
BEGIN
BEGIN TRY
INSERT INTO [CDN].[PracPlanDni]
        ([PPL_PraId]
        ,[PPL_Data]
        ,[PPL_TS_Zal]
        ,[PPL_TS_Mod]
        ,[PPL_OpeModKod]
        ,[PPL_OpeModNazwisko]
        ,[PPL_OpeZalKod]
        ,[PPL_OpeZalNazwisko]
        ,[PPL_Zrodlo]
        ,[PPL_TypDnia])
VALUES
        (@PRI_PraId
        ,@DataInsert
        ,GETDATE()
        ,GETDATE()
        ,'ADMIN'
        ,'Administrator'
        ,'ADMIN'
        ,'Administrator'
        ,0
        ,3)
END TRY
BEGIN CATCH
END CATCH
END

SET @id = (select [cdn].[PracPlanDni].[PPL_PplId] from [cdn].[PracPlanDni] where [cdn].[PracPlanDni].[PPL_Data] = @DataInsert and [cdn].[PracPlanDni].[PPL_PraId] = @PRI_PraId);

-- DODANIE GODZIN W NORMIE
INSERT INTO CDN.PracPlanDniGodz
(PGL_PplId,
PGL_Lp,
PGL_OdGodziny,
PGL_DoGodziny,
PGL_Strefa,
PGL_DzlId,
PGL_PrjId,
PGL_UwagiPlanu)
VALUES
(@id,
1,
DATEADD(MINUTE, 0, @GodzOdDate),
DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate),
2,
1,
1,
'');

-- DODANIE NADGODZIN
IF(@CzasPrzepracowanyInsert > @PracaWgGrafikuInsert)
BEGIN

IF(@Godz_dod_50 > 0)
BEGIN
INSERT INTO CDN.PracPlanDniGodz
	        (PGL_PplId,
	        PGL_Lp,
	        PGL_OdGodziny,
	        PGL_DoGodziny,
	        PGL_Strefa,
	        PGL_DzlId,
	        PGL_PrjId,
	        PGL_UwagiPlanu)
        VALUES
	        (@id,
	        1,
	        DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate),
	        DATEADD(MINUTE, 60 * @Godz_dod_50, DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate)),
	        4,
	        1,
	        1,
	        '');
SET @CzasPrzepracowanyInsert = @CzasPrzepracowanyInsert - @Godz_dod_50;
END

IF(@CzasPrzepracowanyInsert > @PracaWgGrafikuInsert)
BEGIN
INSERT INTO CDN.PracPlanDniGodz
	        (PGL_PplId,
	        PGL_Lp,
	        PGL_OdGodziny,
	        PGL_DoGodziny,
	        PGL_Strefa,
	        PGL_DzlId,
	        PGL_PrjId,
	        PGL_UwagiPlanu)
        VALUES
	        (@id,
	        1,
	        DATEADD(MINUTE, -60 * (@CzasPrzepracowanyInsert - @PracaWgGrafikuInsert), @GodzDoDate),
	        @GodzDoDate,
	        4,
	        1,
	        1,
	        '');
END
END";
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
                                    using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                                    using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                                    using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                            Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                            Console.WriteLine(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                            tran.Rollback();
                            var e = new Exception(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                            e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                            throw e;
                        }
                    }
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine($"Poprawnie dodawno plan z pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                Console.ForegroundColor = ConsoleColor.White;
                tran.Commit();
            }
        }
    }
}