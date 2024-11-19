using All_Readeer;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Wordprocessing;
using ExcelDataReader;
using System.Data;
using static All_Readeer.Grafik_Pracy_Reader;
class Program
{
    private static string Files_Folder = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Pliki pokaz";
    private static string Errors_File_Folder = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Errors\\";
    private static string Bad_Files_Folder = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Bad Files\\";
    private static string Optima_Conection_String = "Server=ITEGER-NT;Database=CDN_Wars_Test_3_;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
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

    public static void Main()
    {
        Check_Foldery();
        //Wpierdol do while(true){} jeśli to tyle
        ZrobToWieszCoNoWieszOCoMiChodzi();
    }

    // TODO MUST now or later
    // Upgrade dodawania zwolnien/urlopów/nieobecnosci
    // co znaczy ob. w grafikach pracy v2 -> dałem nieobecnosc
    // 2 prac o tej samej nazwie
    // prac ktorych nie ma w bazie

    // TODO RACZEJ NIE TRZEBA
    // dodać support dla Zachód - zespół utrzymania czystości - Szczecin - karty pracy.xlsx bo są kurwa w kostke rubika zrobione
    // grafik v2024 SPRAWDZ KILKA GRAFIKOW POD SOBĄ np Centru - Terespol - grafiki ALbo możliwe że nie trzeba
    // TODO Nieobecności w grafik v2024 jeśli takie będą

    // TODO OBY NIE
    // Wyszyścic ten zjebany pierdolony śmierdzący gówno kurwa kod żygać mi się chce
    // TODO fix nieobecnosci przerywane weekendami i nowymi miesiacami itp JEŚLI SIĘ DA a raczej nie i nie no w sumie to nie no raczej nie mhm nie

    public static void ZrobToWieszCoNoWieszOCoMiChodzi()
    {
        string[] filePaths = Directory.GetFiles(Files_Folder);
        if (filePaths.Length == 0) {
            Console.WriteLine("Nie znaleziono żadnych plików");
            return;
        }

        error_logger.Set_Error_File_Path(Errors_File_Folder);
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
                    Console.WriteLine(ex.Message);
                    continue;
                }
            }

            error_logger.Nazwa_Pliku = filePath;

            var (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info(filePath);
            if (Last_Mod_Osoba == "Error") {
                error_logger.New_Custom_Error($"Error reading file {filePath}, could not reed metatata");
            }

            int ilosc_zakladek = 0;
            using (var workbook = new XLWorkbook(filePath))
            {
                Kurwa_Usun_Ukryte_Karty_XD(workbook);
                ilosc_zakladek = workbook.Worksheets.Count;
                for (int i = 1; i <= ilosc_zakladek; i++)
                {
                    error_logger.Nr_Zakladki = i;

                    var zakladka = workbook.Worksheet(i);
                    error_logger.Nazwa_Zakladki = zakladka.Name;
                    var typ_pliku = Kurwa_tego_no_wez_zobacz_ktory_rodzaj_zakladki_to_jest_mordzia_co(zakladka);
                    if (typ_pliku == 0)
                    {
                        error_logger.New_Custom_Error("Nie rozpoznano tego rodzaju zakladki: " + error_logger.Nazwa_Pliku + "nr zakladki: " + error_logger.Nr_Zakladki);
                        continue;
                    }
                    else if (typ_pliku == 1)
                    {
                        try
                        {
                            Karta_Pracy_Reader_v2 karta_Pracy_Reader_V2 = new();
                            karta_Pracy_Reader_V2.Set_Optima_ConnectionString(Optima_Conection_String);
                            karta_Pracy_Reader_V2.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time, 0);
                        }
                        catch
                        {
                            Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                    else if (typ_pliku == 2)
                    {
                        try
                        {
                            Karta_Pracy_Reader_v2 karta_Pracy_Reader_V2 = new();
                            karta_Pracy_Reader_V2.Set_Optima_ConnectionString(Optima_Conection_String);
                            karta_Pracy_Reader_V2.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time, 1);
                        }
                        catch
                        {
                            Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                    else if (typ_pliku == 3)
                    {
                        try
                        {
                            Grafik_Pracy_Reader_v2 grafik_Pracy_Reader_V2 = new();
                            grafik_Pracy_Reader_V2.Set_Optima_ConnectionString(Optima_Conection_String);
                            grafik_Pracy_Reader_V2.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time);
                        }
                        catch
                        {
                            Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                    else if (typ_pliku == 4)
                    {
                        try
                        {
                            Grafik_Pracy_Reader_v2024 grafik_Pracy_Reader_v2024 = new();
                            grafik_Pracy_Reader_v2024.Set_Optima_ConnectionString(Optima_Conection_String);
                            grafik_Pracy_Reader_v2024.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time);
                        }
                        catch
                        {
                            Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                }
            }
        }
    }
    public static int Kurwa_tego_no_wez_zobacz_ktory_rodzaj_zakladki_to_jest_mordzia_co(IXLWorksheet workshit)
    {
        try
        {
            var dane = workshit.Cell(6, 1).GetValue<string>().Trim();
            if (dane == "Dzień") //karta v1
            {
                return 2;
            }
        }
        catch { Console.Write(""); }
        try
        {
            var dane = workshit.Cell(6, 2).GetValue<string>().Trim(); // TODO dodać support dla Zachód - zespół utrzymania czystości - Szczecin - karty pracy.xlsx bo obok siebie i pod są karty xdd
            if (dane == "Dzień") //karta v2
            {
                return 1;
            }
        }
        catch { Console.Write(""); }
        try
        {
            var dane = workshit.Cell(9, 2).GetValue<string>().Trim();
            if (dane == "Dzień") //karta v2 ale jest kurwa niżej xdd
            {
                return 1;
            }
        }
        catch { Console.Write(""); }
        try
        {
            var dane = workshit.Cell(3, 1).GetValue<string>().Trim();
            if (dane.Contains("GRAFIK PRACY")) // grafik v2
            {
                return 3;
            }
        }
        catch { Console.Write(""); }
        try
        {
            var dane = workshit.Cell(1, 1).GetValue<string>().Trim(); // grafik v2024 //TODO SPRAWDZ KILKA GRAFIKOW POD SOBĄ i sprawdz multiples of LICZBA GODZIN
            if (dane.Contains("GRAFIK PRACY"))
            {
                return 4;
            }
        }
        catch { Console.Write(""); }
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
            Console.WriteLine($"File I/O error occurred: {ioEx.Message}");
        }
        catch (UnauthorizedAccessException authEx)
        {
            Console.WriteLine($"Access error occurred: {authEx.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error occurred: {ex.Message} in {newFilePath}");
        }
    }
    private static void Kurwa_Usun_Ukryte_Karty_XD(XLWorkbook workbook)
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
        if (!Directory.Exists(Files_Folder))
        {
            Console.WriteLine("Brak folderu z plikami excel");
            return;
        }

        if (!Directory.Exists(Errors_File_Folder))
        {
            Directory.CreateDirectory(Errors_File_Folder);
        }
        else
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

    }
    private static bool Is_File_Xlsx(string filePath)
    {
        try
        {
            var workbook = new XLWorkbook(filePath);
        }
        catch
        {
            //Console.WriteLine($"Plik to nie arkusz xlsx: {filePath}.");
            return false;
        }
        return true;
    }
    public static void ConvertToXlsx(string inputFilePath, string outputFilePath)
    {
        if (!File.Exists(inputFilePath))
            throw new FileNotFoundException($"Plik {inputFilePath} nie istnieje.");

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