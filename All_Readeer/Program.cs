using All_Readeer;
using ClosedXML.Excel;
using ExcelDataReader;
using System.Data;
using System.Text.Json;

class Program
{
    public static string Optima_Conection_String = "Server=ITEGER-NT;Database=CDN_Wars_Test_3_;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
    public static string Files_Folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files");
    public static string Errors_File_Folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Errors");
    public static string Bad_Files_Folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Bad Files");
    public static string Processed_Files_Folder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Processed Files");
    public static Error_Logger error_logger = new();
    public static string sqlQueryInsertObecnościDoOptimy = @"
DECLARE @id int;

                -- dodawaina pracownika do pracx i init pracpracdni
IF((select DISTINCT COUNT(PRI_PraId) from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1) > 1)
BEGIN
	DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników w bazie o takim samym imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
	THROW 50001, @ErrorMessageC, 1;
END
DECLARE @PRI_PraId INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRI_PraId IS NULL
BEGIN
	SET @PRI_PraId = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRI_PraId IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie o imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
		THROW 50000, @ErrorMessage, 1;
	END
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
                    ,@DataMod
                    ,@DataMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,@ImieMod
                    ,@NazwiskoMod
                    ,0)
    END TRY
    BEGIN CATCH
    END CATCH
END

SET @id = (select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId);

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
		@TypPracy,
		1,
		1,
		'',
		1);";
    public static string sqlQueryInsertNieObecnoŚciDoOptimy = @$"
IF((select DISTINCT COUNT(PRI_PraId) from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1) > 1)
BEGIN
	DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników w bazie o takim samym imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
	THROW 50001, @ErrorMessageC, 1;
END
DECLARE @PRACID INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRACID IS NULL
BEGIN
	SET @PRACID = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRACID IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie o imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
		THROW 50000, @ErrorMessage, 1;
	END
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
    public static string sqlQueryInsertPlanDoOptimy = $@"
DECLARE @id int;
IF((select DISTINCT COUNT(PRI_PraId) from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1) > 1)
BEGIN
	DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników w bazie o takim samym imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
	THROW 50001, @ErrorMessageC, 1;
END
DECLARE @PRI_PraId INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRI_PraId IS NULL
BEGIN
	SET @PRI_PraId = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRI_PraId IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie o imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
		THROW 50000, @ErrorMessage, 1;
	END
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
        ,@DataMod
        ,@DataMod
        ,@ImieMod
        ,@NazwiskoMod
        ,@ImieMod
        ,@NazwiskoMod
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
    public static void Main()
    {

        error_logger.Set_Error_File_Path(Errors_File_Folder);
        Check_Foldery();
        GetConfigFromFile();
        error_logger.Set_Error_File_Path(Errors_File_Folder);
        while (true)
        {
            Thread.Sleep(3000);
            Do_The_Thing();
        }
    }
    /// <section>
    // TODO MUST now or later
    // Upgrade dodawania zwolnien/urlopów/nieobecnosci z grafików v2 bo do nich nie mam w sumie ządnego kodu wiec wszystko co jest nierozpoznane daje jako nieobecność
    // co znaczy ob. w grafikach pracy v2 -> dałem nieobecnosc
    // 2 prac o tej samej nazwie
    // prac ktorych nie ma w bazie

    // TODO RACZEJ NIE TRZEBA
    // Nieobecności w grafik v2024 jeśli takie będą
    // Lepsze rozpoznawanie typów grafików
    // TODO OBY NIE
    // Wyszyścic ten zjebany pierdolony śmierdzący gówno kurwa kod żygać mi się chce
    // fix nieobecnosci przerywane weekendami i nowymi miesiacami itp JEŚLI SIĘ DA a raczej nie i nie no w sumie to nie no raczej nie mhm nie

    //UWAGI:
    // Ja na ten moment tak sobie zrobiłem:
    /*RodzajNieobecnosci.UO => "Urlop okolicznościowy",
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
    _ => "Nieobecność (B2B)"*/
    // Przyczyny:
    /*
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
    _ => 9                             // Nie dotyczy dla pozostałych przypadków*/

    // co znaczy ob. w grafikach pracy z przed 2024 -> dałem nieobecnosc B2B na ten moment
    // czy UŻ powinienem traktować jako UZ (Urlop na żądanie) czy coś innego
    // Pracownicy którzy mają te same imie i nazwisko -> ciężko mi je połączyć gdyż często dział/zespół/stanowisko czy inne nie zleją sie dokładnie z danymi w bazie
    // Mam wielu pracowników których nie ma w bazie ciężko rozróżnić czy nie ma go w bazie czy jest źle wpisany
    //
    /// </section>
    public static void Do_The_Thing()
    {
        string[] filePaths = Directory.GetFiles(Files_Folder);
        if (filePaths.Length == 0) {
            Console.Clear();
            Console.WriteLine($"Nie znaleziono żadnych plików w folderze {Files_Folder}");
            return;
        }
        foreach (string current_filePath in filePaths)
        {
            string filePath = current_filePath;
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Czytanie: {System.IO.Path.GetFileNameWithoutExtension(filePath)}");
            Console.ForegroundColor = ConsoleColor.White;
            if (!Is_File_Xlsx(filePath))
            {
                try
                {
                    ConvertToXlsx(filePath, Path.ChangeExtension(filePath, ".xlsx"));
                    filePath = Path.ChangeExtension(filePath, ".xlsx");
                }
                catch (Exception ex)
                {
                    //Console.WriteLine(ex.Message);
                    MoveFile(filePath);
                    continue;
                }
            }

            error_logger.Nazwa_Pliku = filePath;

            var (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info(filePath);
            if (Last_Mod_Osoba == "Error") {
                error_logger.New_Custom_Error($"Błąd w czytaniu {filePath}, nie można wczytać metadanych");
            }
            error_logger.Last_Mod_Osoba = Last_Mod_Osoba;
            error_logger.Last_Mod_Time = Last_Mod_Time;

            int ilosc_zakladek = 0;
            using (var workbook = new XLWorkbook(filePath))
            {
                Usun_Ukryte_Karty(workbook);
                ilosc_zakladek = workbook.Worksheets.Count;
                for (int i = 1; i <= ilosc_zakladek; i++)
                {
                    error_logger.Nr_Zakladki = i;
                    var zakladka = workbook.Worksheet(i);
                    error_logger.Nazwa_Zakladki = zakladka.Name;

                    var typ_pliku = Get_Typ_Zakladki(zakladka);
                    if (typ_pliku == 0)
                    {
                        Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                        error_logger.New_Custom_Error("Nie rozpoznano tego rodzaju zakladki: " + error_logger.Nazwa_Pliku + " nr zakladki: " + error_logger.Nr_Zakladki + " nazwa zakladki: " + error_logger.Nazwa_Zakladki + " Porada: Sprawdź czy nagłówki są uzupełnione");
                        Console.WriteLine("Nie rozpoznano tego rodzaju zakladki: " + error_logger.Nazwa_Pliku + " nr zakladki: " + error_logger.Nr_Zakladki + " nazwa zakladki: " + error_logger.Nazwa_Zakladki + " Porada: Sprawdź czy nagłówki są uzupełnione");
                        continue;
                    }
                    else if (typ_pliku == 1)
                    {
                        try
                        {
                            Karta_Pracy_Reader_v2.Process_Zakladka_For_Optima(zakladka);
                        }
                        catch
                        {
                            try
                            {
                                Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            }
                            catch { }

                            continue;
                        }
                    }
                    else if (typ_pliku == 2)
                    {
                        try
                        {
                            Grafik_Pracy_Reader_v2.Process_Zakladka_For_Optima(zakladka);
                        }
                        catch
                        {
                            try
                            {
                                Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            }
                            catch { }
                            continue;
                        }
                    }
                    else if (typ_pliku == 3)
                    {
                        try
                        {
                            Grafik_Pracy_Reader_v2024.Process_Zakladka_For_Optima(zakladka);
                        }
                        catch
                        {
                            try
                            {
                                Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            }
                            catch { }
                            continue;
                        }
                    }
                }
            }
            MoveFile(filePath);
        }
    }

    private static void MoveFile(string filePath)
    {
        try
        {
            string destinationPath = Path.Combine(Processed_Files_Folder, Path.GetFileName(filePath));

            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }

            File.Move(filePath, destinationPath);
        }
        catch (Exception e)
        {
            //error_logger.New_Custom_Error(e.Message + " Plik: " + filePath);
        }
    }
    private static int Get_Typ_Zakladki(IXLWorksheet workshit)
    {
        foreach (var cell in workshit.CellsUsed()) // karta v2
        {
            try
            {
                var cellValue = cell.GetString();
                if (cellValue.Contains("Dzień"))
                {
                    return 1;
                }
            }
            catch { }
        }

        var cellValue3_1 = workshit.Cell(3, 1).Value.ToString();
        if (cellValue3_1.Trim().Contains("GRAFIK PRACY")) // grafik v2
        {
            return 2;
        }

        var cellValue1_1 = workshit.Cell(1, 1).Value.ToString();
        if (cellValue1_1.Trim().Contains("GRAFIK PRACY")) // grafik v2024
        {
            return 3;
        }

        return 0;
    }
    private static (string, DateTime) Get_File_Meta_Info(string File_Path)
    {
        try
        {
            using (var workbook = new XLWorkbook(File_Path))
            {
                DateTime lastWriteTime = File.GetLastWriteTime(File_Path);

                if (workbook.Properties.LastModifiedBy == null)
                {
                    return ("", lastWriteTime);
                }
                return (workbook.Properties.LastModifiedBy, lastWriteTime);
            }
        }
        catch
        {
            DateTime lastWriteTime = File.GetLastWriteTime(File_Path);
            return ("", lastWriteTime);

        }
    }
    private static void Copy_Bad_Sheet_To_Files_Folder(string filePath, int sheetIndex)
    {
        var newFilePath = System.IO.Path.Combine(Bad_Files_Folder, "copy_" + System.IO.Path.GetFileName(filePath));
        try
        {
            using (var originalwb = new XLWorkbook(filePath))
            {
                var sheetToCopy = originalwb.Worksheet(sheetIndex);
                string newSheetName = $"Copy_{sheetIndex}_{sheetToCopy.Name}";
                if (newSheetName.Length > 31)
                {
                    newSheetName = newSheetName.Substring(0, 31);
                }
                using (var workbook = File.Exists(newFilePath) ? new XLWorkbook(newFilePath) : new XLWorkbook())
                {
                    if (workbook.Worksheets.Contains(newSheetName))
                    {
                        return;
                    }
                    sheetToCopy.CopyTo(workbook, newSheetName);
                    var properties = originalwb.Properties;
                    properties.Author = "Copied by program";
                    properties.Modified = DateTime.Now;
                    workbook.SaveAs(newFilePath);
                }
            }
        }
        catch (IOException ioEx)
        {
            //Console.WriteLine($"File I/O error occurred: {ioEx.Message}");
        }
        catch (UnauthorizedAccessException authEx)
        {
            //Console.WriteLine($"Access error occurred: {authEx.Message}");
        }
        catch (Exception ex)
        {
            //Console.WriteLine($"Error occurred: {ex.Message} in {newFilePath}");
        }
    }
    private static void Usun_Ukryte_Karty(XLWorkbook workbook)
    {
        var hiddenSheets = new List<IXLWorksheet>();
        foreach (var sheet in workbook.Worksheets)
        {
            if (sheet.Visibility == XLWorksheetVisibility.Hidden)
            {
                hiddenSheets.Add(sheet);
            }
        }
        foreach (var sheet in hiddenSheets)
        {
            workbook.Worksheets.Delete(sheet.Name);
        }
        workbook.Save();
    }
    private static void Check_Foldery()
    {
        if (!File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json")))
        {
            File.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json")).Dispose();
            var defaultConfig = new
            {
                Files_Folder,
                Errors_File_Folder,
                Bad_Files_Folder,
                Processed_Files_Folder,
                Optima_Conection_String
            };
            File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json"), JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
        }

        if (!Directory.Exists(Files_Folder))
        {
            Directory.CreateDirectory(Files_Folder);
        }

        if (!Directory.Exists(Errors_File_Folder))
        {
            Directory.CreateDirectory(Errors_File_Folder);
        }


        if (File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Errors.txt")))
        {
            File.WriteAllText(Errors_File_Folder + "Errors.txt", string.Empty);

        }

        if (!Directory.Exists(Bad_Files_Folder))
        {
            Directory.CreateDirectory(Bad_Files_Folder);
        }
        else
        {
            foreach (var file in Directory.GetFiles(Bad_Files_Folder))
            {
                File.Delete(file);
            }
            foreach (var directory in Directory.GetDirectories(Bad_Files_Folder))
            {
                Directory.Delete(directory, recursive: true);
            }
        }

        if (!Directory.Exists(Processed_Files_Folder))
        {
            Directory.CreateDirectory(Processed_Files_Folder);
        }
        else
        {
            foreach (var file in Directory.GetFiles(Processed_Files_Folder))
            {
                File.Delete(file);
            }
            foreach (var directory in Directory.GetDirectories(Processed_Files_Folder))
            {
                Directory.Delete(directory, recursive: true);
            }
        }

    }
    private static bool Is_File_Xlsx(string filePath)
    {
        try
        {
            var workbook = new XLWorkbook(filePath);
        }
        catch
        {
            //Plik to nie arkusz xlsx
            return false;
        }
        return true;
    }
    public static void ConvertToXlsx(string inputFilePath, string outputFilePath)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        DataSet dataSet;
        using (var stream = File.Open(inputFilePath, FileMode.Open, FileAccess.Read))
        using (var reader = ExcelReaderFactory.CreateReader(stream))
        {
            var config = new ExcelDataSetConfiguration
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            dataSet = reader.AsDataSet(config);
        }

        using var workbook = new XLWorkbook();
        foreach (DataTable table in dataSet.Tables)
        {
            var worksheet = workbook.Worksheets.Add(table.TableName);
            for (int i = 0; i < table.Columns.Count; i++)
                worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    var value = table.Rows[i][j];

                    if (value == DBNull.Value)
                    {
                        worksheet.Cell(i + 2, j + 1).Value = string.Empty;
                    }
                    else
                    {
                        worksheet.Cell(i + 2, j + 1).Value = value.ToString();
                    }
                }
            }
        }
        workbook.SaveAs(outputFilePath);
        var (o, d) = Get_File_Meta_Info(inputFilePath);
        workbook.Properties.LastModifiedBy = o;
        workbook.Properties.Modified = d;
        workbook.SaveAs(outputFilePath);
    }
    public static void GetConfigFromFile()
    {
        var filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json");
        if (!File.Exists(filePath))
        {
            File.Create(filePath).Dispose();
            var defaultConfig = new
            {
                Files_Folder,
                Errors_File_Folder,
                Bad_Files_Folder,
                Processed_Files_Folder,
                Optima_Conection_String
            };
            File.WriteAllText(filePath, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
        }
        string json = File.ReadAllText(filePath);
        var config = JsonSerializer.Deserialize<Config_Data_Json>(json);
        if (config != null)
        {
            Files_Folder = config.Files_Folder;
            Errors_File_Folder = config.Errors_File_Folder;
            Bad_Files_Folder = config.Bad_Files_Folder;
            Processed_Files_Folder = config.Processed_Files_Folder;
            Optima_Conection_String = config.Optima_Conection_String;
        }
    }

}
public class Config_Data_Json
{
    public string Files_Folder { get; set; } = Program.Files_Folder;
    public string Errors_File_Folder { get; set; } = Program.Errors_File_Folder;
    public string Bad_Files_Folder { get; set; } = Program.Bad_Files_Folder;
    public string Processed_Files_Folder { get; set; } = Program.Processed_Files_Folder;
    public string Optima_Conection_String { get; set; } = Program.Optima_Conection_String;
}

//{
//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::.................................::::::::----::::::-==+=======--:----==
//::::::--::::::::::::::::::::::::::::::::::::::::::::::::::::::...............::::..................:::::-------:::::===+====---:::::---
//:::::::=+-::::::::::::::::::::::::::::::::::::::::::::::::::.........:::::::::::::...............:::::::::::----:::--======----------:-
//::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::::.......:::::::::::::::.........::::::.........::----:-----======---------
//:::::::::::::::::::::::::::::::::::::::::::::::::::::::::.::::::..........................:--::::::...........:---------=-=========---=
//::::::::::::::::::::::+++=::::::::::::::::::::.:::::..:.................................::----::::::::::::::::::-------=-==============
//:::::::::::::::::::::::::+-.:::--=-:::::::::.::.:......................................:-====------::::::::::::----===-============++++
//::::::::::::::::::::::--=++**=-:::::::.:...:::..:.......................................-===-------:------::::---==========-====+++++++
//....::..:::..:::::::::.:::--+*-::::.:..::..::.:................:::::::.................:--====-------===---::---===++=+===:::-===++++++
//.........:........:.:::::::::.:...................:::::::.......::::..................:---======----==+=+==--:===++++++=-::::-----=====
//................................................:::::::.................:.............:---==+====-====+++++==--+++++++==--::::::-------
//...............................................:--:::-:::.............................:--=+++==========++**+==+++++++===---:::.......::
//...............................................:-:::::::...............................:=+***+=======-========++++==----::::::.........
//..............................................::::::::::...............................:-+***++=====-=====---:::-==--:::.....:.........
//............................................:::::::::::.....................+**+-.......:=++*+++==-----=====--::---......-**=-=--......
//............................................::.:::::::::::.................=#####+:......:-=++++=----:::-=-=-------:...:++#%==+#=-.....
//.............................................:::::::::.....................#**#######*+++:.:====-:::.:..::-==----------=+=##+=+#+=::...
//.............................................:::::::::....................+#***##*##%%%%%##*+=--::..::::.::-===-:----::=+=+++=**--=+===
//............................................::::::::.....................*+=#=+*#**%%%%#####+==------:::.:::----:===--:-#=**+*#=-++==+-
//........................................:::.::::::.......................=*++=++**+*#####*%#*+==-==+====:::::===-====-=##*+*##++***+***
//..............................................:::::....................**#+#++==++**%#*###%##+++=+++++*+=---:-===+++++=*###%%+##+##****
//........................................:......:......................=+#*#*+**#+##%%#*###%%%%#**+========--==+++**#*+*#####**%#+###*##
//......................................:::........................=+:.:#####%*=*+*=+**##**##%%+++==-:::--=======++*##*##***###%#=####%%#
//...................................:::::::::::................+*****#####+%#++++*#*##*#**#%#=-=---=++++=+=====--==+**#+:.::=###**##%%**
//....................................:::++++**==.............=**=+===+#+##+#*+*#**###***#*###=----=+**#**++=---:::::::.:.....:##**%##%**
//...............................::....-#+*#%%#*=++:......==+=+*++*+*+=****##*##+++*#****##***+-===+*####**+=---::::::........:=%%%@%#*#+
//..........................::--==:::::++-+=##*++==+.....=*+++*%+=#*#*+*#*+###***%****#*###**=----=+******++==-------:::::....:-*%%%%##*+
//....................:::----===----=:==+-++#*+#**++=...##%*+*%#=*++**+%%#***#++*=****+######*+=-==+*****+===---:---=-----::::::=#%%%#+#*
//............::.......::--====--::::-+++**##*#++++*##::***=*+#*+++=++*%#####*##+*+**#%##%##*#*+==+*****++=--:::..:--=====----:-+#*%#****
//...........:-:::---:--=-----:::....+**++********%#*###*#*%===+==+=++=+#######++**#%#########**++++****++=--::::..:--====-----=+**##*+**
//::::::::::--::----==-==-:...:....-++++*****+++**%#*%#*++##+++*#+*++**+*######+#**######%##*##**+++*****+=-----::::--------:--=*##***###
//:::::::::-:-----====-:::::::....:=**+++**+#*+++#####***+*++**#*=+*+===+#%###**#######%%%########*******+++++++=-::::::..:::-=+**#*++***
//:::::::::::----====-::-----:....-+*#**+++*+++*+#####*##+#++=+**=****+*+*###*####*##*####%%%#%%%%###*+*++++++++==--:......::-=**##**#+-+
//:.::::::::----=====--=====-:=**+#***++++*#***#+#*****#+#*##=*====++++++##*#######*+**#%##%%%#%###%##*=+***+++*+++==--...:::-=*###%%%*++
//.::::::::-**+==++=====----++*++***+++==**#*=+*#%*#*++*+*%%*#**=++++=+++#+#######*=+==++**###*#%######************+++=-.:::=+*##%#%#*+*#
//..:::::::=++++=+=+++*#=:--**#++*+==+=-++*****+***+*+*++*#++=*#%*#**++++***+:.................:+%%###+=***#######**+++=-::=+***###**#***
//..::..::::-=---===++*##-=#%+++=++-==**===-=++++++++++=*#*--+*#*##*****+.....:-=-+***####*#+:......+=:::-+########**:.-=-=+**+*###*#*+++
//....::::....::::-===++***#*+==+++==++-:-+==+***####*++##%*=****++*#-....:-+##%%%%%%%%%%%%%%#+=---:...---:--*%#%##*+=-==-=+++*+******=+*
//%%+:::...........-===---**+==*+=++-=-=*++===#+***+===+=+#=**#++**....:=*#%%%%%%%%%%%%%%%#%@@@%*+*##*-..=---+#####*+::==+**++**#*###+++*
//%#-:.............:-::..=*%+=+=*+=+=+++===+++==+++=++++*####***+...:-+#%%%%%%%%%%%%%%%%%%%%@@@%#*#%%%%#+::++**####*++=:.+**+****##%***##
//##=:.....:.......:::....=+=-==+#=-=+==:-*+=+*+=++#*++=*+**+*#+..:+#%%%%%%%@%%%%%%%%%%%@%@@@@%###%%%%%%%#+:+*#####**++=.:-+***%#%%%%%#+*
//**=:...::--::.............:::-=+*+===-+++*+*++++*===++==+++*=.:+#%%%%%%%%%%%%%%%%%%%@%%%%%%%%%%@%%@@@%%%%#=*#%##***+=-.:-=+#######%####
//**+=--:-----:--:..........:::-*+*=-=-==++++=*===++-+++*+==+=.=*#%%%%%%%%%%%%%%%%%%%%%%%%@%@@@@@%@@@@@@%%%%#*#####*+=::--=+*####*###*##+
//+++++==========-:..........:-=*=-=-=+=++**+=+==+++++****==+:+*#%%@%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%@%%%%%%####**==-::-=******+*+*+++=
//=+++++===+======-:.........:=::::--=++++++==**+++**##*****=*#%%%%%%%@%@%%%%%%#%%%%%%%%%#####%%%%%%%%%%%%%%%%%##*+====-==+++===++++***++
//--+**+***+====++=-:........:.....:-========+===+*******++=*#%%%@%@%%%%%%%%%###%###########%##%####%%#%%%%%%%%#++==============-----==-=
//::-===+*++=====**=-::::...........::-=========-=++*+++==+=#%%@@@@@%%%%%%%%#########%###################%##%%%#===-----------::.::::::-:
//::--:-=--::::--++=---:::.......:::::--==--=------=====--=-#%@@@@@%%%%%%%%####*###########**********+++**######*=-----------:::::::::::.
//:::-::-:-..::-=++==-:::..::::.:::::---===------===-=====--#@@@@@@@@%%%%%%########**+++*%%#*****###**++*#%%%#%%@*===---++=:::::::::..=-.
//:::--=+-:-..:-=+++==-::.......:::----=------===========---#%@@@%%@%%%%%%%#####*+*#%*++***+===+#@@@@@@@@@@@@@@@@#%+*=+=**+===-:::::-=+:.
//-==*++*+=---=***++=---::::.....:---======-=========+#+-:::#@@@@%@%%%%%%#####%#****#+=*#%@@@@@%%%%%%%%@@@@@@@@@@@@@***+*++=+++=-:..-=-::
//=++*#***+=+******++=-----::....:---=-===========+--*=--:-:#@@@@@%%%%%###%#+###-+@%%%%@%%%%@%%%#%##%#%%@@@@@@@@@@@@%#***++++=-::::::+--+
//*++*##***+*******++++==-=+=-:::----====+============++--::=@@@@%%%##%%@%**#@@@@@%%%%%%%##%%%%@@@@@@@%%%%@@@@@@@@@@@@#+=++=-=--==:-=+=++
//*##*##**##*****##++*+**++=---::----=+=+==++===+**+++*+:::::%@@%%#%@@**%%%#####%@@@@@%%#*%@@@%#+++++*%@%%@%@@@@@@@@@@%===-====+==:.:*=++
//******#*##+******+++*++----::::------==-=+=--=+++-+==-:::..+@@%%@#%@@%#*%@#@@@#@@@@@%%+===#@@*#@@*@#+=#%@@@@@@@@@@@@#-=+-=++==+=::-*+==
//%#%*@*#*****#**#*=-+*==-----------===+---=---======:+:.....-%#*#@@@@%#*@%#*@@#=##%%@%%###%#%#+#@%*+@%*+#@@@@@@@@@@@%=-+---=*--+:-=-=---
//%#@%#******++**++------#=::.:##*=---+==++-=====+=-=+-:-....-*%@@@@@@%%%%%%#**#+-+%@@%%##%%%#=++-=#*#%%%#%@%%@@@@@@@*-==-==:=*=--===+=-+
//###**#*#**#####++#*=-::-+-::-=%#*-=***+---:-=+++=-+=:::....#%@@@@@@@#%%%%%#*+++#%@%@%**##%%%*+====**###%%@@@@@@@@@*--=--=--=:#*-=-+*+==
//###*##*#*###%##*+++=-:--+=+***#%##===++++#%%*=++=-+=:-::::*@@@@%%@@%%##***###%%%%%%%%**##%%%%%*+*+++*+*#%@@%%@@@@%+==-===--+-+*+:-*+===
//##*++**#*######**++==::-+++*%%%%@###*%%%#%#*++++++====-::-@%@@@%@@@%%%#*#*##%###%%%%%***#%%#%##%#**+***#%%@@%@#%@%==-=+-====-=-#*#*--=-
//#++*+***#######**+++=:--=++*####%#####=--====+=*+=++===--*%%@@@%%%@@%%%####%%###%@%%%###%%#%%#######**%%@%%%*:-*%*--:-=-=--===:*#*=:---
//###%#%%######***++===::-+######%@%#--=--======++++*++==-=*%%@@@@@%@@@%%%%%%####*#%%*===+#%%***###%%%%%%@%@%#-:-=@+:--:=-+--==-:+*+::*+=
//%%%%%%%%%##******#**+=+++++*#%%%#::=-+====-=-=+++=+**+=---%%@@@@@%@%@%@%%%%#**++#*#%*==*###*+++*#%%%%%%@@%%=::-:%*--==+:=---=::**+:-*+:
//+*++**=*###***##%%%*%##%%%+=+%#:-=---*---==-==++==*+====::+%@@@@@@%%%%%@%%##****%#%@%##@@%%##***##%%%%%#%%=::--:#+---*=:+-=---:*==::*=:
//+##%-.#==#***#*#%%%*#*#%%#+---==:--+++-====-+=*===++=-+:::.+%@@@@@@%%%%@@%##***%%%%%@%%@@%%%%****#%@%@%%*#-::=-:#+:=-+--=-==:=-*++-:#*-
//*#%%%#%#=#####*#%%##*=*###+:+%%+==*%@*+==:==+=*===+=-=+:-::+%@@%@@@@%%%@@%####%%%%%%####%%%%%%#####%%@%%%:--:-=-#+-=-+-=+===-=-**=--*+-
//#%@@@%%%###*#%#%%%#**+***#+:##%#++###+*+=+=--#+===*+==+-=::-%%:::.:::+%@@%%%%%%@%%*++**###%@@%%%###%%*::-:--:==:*+--=-=:+=+===+**+:=*++
//++++++*+++**#*#*=#*###%%%%%%%%%%%%%%%%#%%%%%%%#+==*==++==---+--:--:::.-%@%%%%%%#%@@@@@@@@@%%%%%%%%%%%+::--++-==-==--+-+-=-===+#*++==*+-
//###%%%%%%%%%@@@@@@@%%%%%%%%%%%%%%%%%%%%%%%#####=++*==++=++=-=--:=::::..#@%%%%%@@%%%#*#####%@@%%%%%%%%=:---+*+==-=+===+++==++-+#+++=#+++
//###%%%%%%%%%@@%%%%%@%%%%%%@%%%%%%%%%%%@%%%@@@%**+*#=***++==:==-:=:::::.=%%%%%@@@%%%@@@@%%@%%@@%%%%%%%:::-=*+++==++==*+#+=++*=+#+++*#++#
//*#%%%%%%%%%%%%%%%#%%%%%%#%##%####%%%%%@@%%@%%#++++*++**==#*+-+#-+:::::..%%%%%%@@@@@@@@@@@@%%@@@%%%@%+:::-=#==+==#*+=+**+++++-#*+==**=+*
//####%%%%%##%%%%%#%%%%%##%##########%%%@%@@@@@@*++*******##++%*#-=-=--:..%@%%%%%%%###*#*#%#%%%%%%%@@@+::==+*=++=+#=+++**=:-*=+#*+++*#+==
//####%%%%##%####%###%%%#%##*###%#####%%#%%%%%%#*+-=#++:-+#+=+%%#++-=+=--:%@@%%%%#%######%###%%%%%%@@%#*:==#+:===*%++=++*++**==%*+#*##+#*
//####%%###%#######%#%%%#######+***#%@@@@@%@%@@%+****=-++++-=#**#=***+++:=%%@@@%@%%%%%%%%%%%%%%%%@@%%%@#=*.:+-**+#+=++=+=-=##=*%++*#%#**+
//%####%####%%##*##*#%#%###****#+%@@@%@@@@@@@@@@%*++=+*#+++-==+#++==+*-:%@%%%@@@@@%@%@%%%%@%%%@@@%%%%%@%#*-++=*##%%%%%%#%##**#%%*+###*##+
//%%%%%%@%%%%%%##%%%%%%%%%%%%%%@@@@%%%%@@@%#####===+***#+#=+#%%%%##%=.:%%#%%%%@@@@@@@@@@@@@@@@@%%%%%%%@@%%#@*.-*%%%%%%#****%%%%%#++#%%#*+
//%%%%%%%%%%%#%%%%%%%%%%%%%%%%%@@@%%%%@@%##@@##%*=--#*##*#*=*#+#+#*..*%%@#%%%%%%%%%@@%@%@@@%%%%%%%%%%%@@@@@%%#+.-###**+***=#%@%*#*+=#%##*
//%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%@%%@@@@@@@@%@=-=+#*#**%*:+#+%%=:.:##@*+%%%%%%%%%%%%%%#+***#%%%##%%@@@%#%%%%%=#%*==+=+**=*%%%**+#====*#
//%#%#%%##%%#%#%%%#%%%%%%%%%%@%%@@%@%%#@%%#%%%%%*:=+%*##*#=-:+*::..+%#%@#%%%%%%%%%%%##*+*****######%@@@%%%+##%*@#++*###*#*++**%*####*#*==
//%#####%#%%%%%##%#*+-#%%%%%%%%%%%@%%%@@@%*%@#+%**+**+%#*#++=.:=:.-%+*%@%%%%%%#%%%%%%#*****#######%%%%@@%%%#=+#%@%=:.=@#%%%%%%**##%%%###*
//#####%%%%%%%%#%*##++*##%%%%%%%%%%%%%@@@@%%@@@%*##**##+*#+:-*@+.+%=#%@@%@%%%%%%%%%%%%%%##########%%%@@@#%@%#%@%%%%%-:=--*%%%%@%%%%%#%%%#
//###%%%%%%#%%%#%%%#*+---==#%%%%%%%%%#@@@#*%%%#+#**++*#=....*#=.*##*##@%%%%%%%%%%%%%%%%##########%%%%%%@####-##%%##@#--*=--#%%@%%@%#%**##
//%%#%%%%%%%%%%##%*+=:=-+*:-:*##%%%%%%@@@%%%@%%%*+*%+...:=+##-**++#%**@@%@%@%%%%%%%%%%%%########%%%%%%@@%%#%@*-####@%*-*@#%*+==*=+%@%%@%%
//%%%%%%%%%%%%%%%%%#**===+==-:-#%%%%%%@%%%%%%@@%%#=::-+#%###=-=++=*###%@@%@@%%%%%%%%%%%%##%%####%%%%%%%@#%%%%###%*#%@%*=*%%#*-#%@%@%+@%%@
//%%%%%######%%%%*##=:-+-#%+=-=*######@%%@@@%%%-:::+%##%%%%+#@@%#%*#*#%%%#%@%%%%%%%%%%%%%%%%###%%%%##%%@*%%%%*%%***#@%%+=*%%#=*##%%%@#%%%
//%%%%#=+###***%#*--#+***%%*==:+%#%%#*+#%%%%*=-=****++###%%*#%*@@@%#*#@%#%#%%%%%%#*%%%%%%%#+*#####*%#*#@##*%*#*%#*+=%#%%**#%%#::::=-==-#%
//#***+#++#++**++:+##=+*+*%%#+=-#**+*%%#%*++=+++#%@@%#*%%%##%###%#@%@%%***#*##***+==**###*++*++++==+=++%#%*%=%+##%##%#%%#####**+=####*#+-
//%%#+=*+**#%%*===+#%===##%+-:=:#+++%%*=--+###+*+*%@@%%%%%#######*%*##%######***+++==++************###%%#**#+@%###%%%%@@#%%#*%**==%%+####
//#+#*#++*%#%%#-=-#-*+=*%#=-=+++#+#%*+++**###+***#%%%%%##%*###%@%#=@**#%#####*****+++++**###*#***#*###%@#+*=#@%###%%%%%%#%%%#*+%++#%##**#
//*-#%#=*#*+%##=:-*=-==%*-:=*+#%+#*##%%*%*#%##@%*%#%@%#%##@#####%#+##-########++**+++++*****######%##%%%*##-%%%###%%%%%%#%%*##*@%+#%###*#
//%+####*#*###+--=-=+##=-:%%*:#%#=%%%#%###*%#%#%%#%%%%@%##%%%#*#%@%+%=#####***#++**++**#####***##*###%%%*#++%%%%*%%%%%%%##*#####@#%%%%@#%
//*%*=***+***+++=++=%%*+=%%##%%#+-##*%%##%#%*+#%%#%%@%%%##@@%%++%@#+***####***+#*++++******#****#####%%%#*=%@%%@**%@%%%%##%%##**%%#@%%%%%
//*#%####+**=::+*-#####%+%%%%%#+=#%###@%#@%#***%@%%%%%@@%%*@@#**+#@**%*###*****+**+++*****+****######%@%*#*@@%%%@#%@@%@%##%%%##*####%@%%%
//*#%%%%**%%**+=*#%#+=+#%%%%%#***#%*#%%%%%%#*#*#%@%#%%%%@%##@%+*%+%#*%*#******+++***+****+*+++*#*####%@%**@@@%%%%##%%@%%##%%##%#*##%%@%@@
//+##%%%%+=#%%###%%#*+#%#%%%#*+*##%#%@@##%%@%*#%%%%##%%@%%#%%@%+%*#**%#%##*******+**+****++*++****###%%#*%@@@%%%###%#%%%++#%%%####%%#%@%%
//}