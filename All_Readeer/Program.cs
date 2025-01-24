using All_Readeer;
using ClosedXML.Excel;
using ExcelDataReader;
using System.Data;
using System.Text.Json;

class Program
{
    public static bool Clear_Logs_On_Program_Restart = false;
    public static bool Clear_Processed_Files_On_Restart = true;
    public static bool Clear_Bad_Files_On_Restart = true;
    public static string Optima_Conection_String = "Server=ITEGERNT;Database=CDN_Wars_prod_ITEGER;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
    public static List<string> Files_Folders = [];
    public static Error_Logger error_logger = new();
    public static readonly string sqlQueryGetPRI_PraId = @"
-- weź @PRA_PraId z akronimu
IF @Akronim IS NOT NULL AND @Akronim != 0
BEGIN
	DECLARE @AkroRes INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @Akronim);
	IF @AkroRes IS NOT NULL
	BEGIN
		SELECT @AkroRes;
	END
END

-- weż @PRA_PraId z imie i nazwisko
IF (
    (
        SELECT DISTINCT COUNT(PRI_PraId)
        FROM cdn.Pracidx
        WHERE
            (PRI_Imie1 = @PracownikImieInsert AND PRI_Nazwisko = @PracownikNazwiskoInsert AND PRI_Typ = 1)
            OR
            (PRI_Imie1 = @PracownikNazwiskoInsert AND PRI_Nazwisko = @PracownikImieInsert AND PRI_Typ = 1)
    ) > 1
)
BEGIN
    DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników w bazie o takim samym imieniu i nazwisku, a takiego akronimu nie ma w bazie: ' + @PracownikImieInsert + ' ' + @PracownikNazwiskoInsert + ' ' + Convert(VARCHAR(200), @Akronim);
    THROW 50001, @ErrorMessageC, 1;
END

DECLARE @PRI_PraId INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRI_PraId IS NULL
BEGIN
	SET @PRI_PraId = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRI_PraId IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie o imieniu, nazwisku i akronimie: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert + ' ' + Convert(VARCHAR(200), @Akronim);
		THROW 50003, @ErrorMessage, 1;
	END
END

DECLARE @EXISTSPRACTEST INT = (SELECT PracKod.PRA_PraId FROM CDN.PracKod where PRA_Kod = @PRI_PraId)

IF @EXISTSPRACTEST IS NULL
BEGIN
    INSERT INTO [CDN].[PracKod] ([PRA_Kod] ,[PRA_Archiwalny],[PRA_Nadrzedny],[PRA_EPEmail],[PRA_EPTelefon],[PRA_EPNrPokoju],[PRA_EPDostep],[PRA_HasloDoWydrukow])
    VALUES (@PRI_PraId,0,0,'','','',0,'');
END
SELECT @PRI_PraId;";
    public static readonly string sqlQueryInsertObecnościDoOptimy = @"
DECLARE @EXISTSDZIEN DATETIME = (SELECT PracPracaDni.PPR_Data FROM cdn.PracPracaDni WHERE PPR_PraId = @PRI_PraId and PPR_Data = @DataInsert)
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
		((select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId),
		1,
		@GodzOdDate,
		@GodzDoDate,
		@TypPracy,
		1,
		1,
		'',
		1);";
    public static readonly string sqlQueryInsertNieObecnoŚciDoOptimy = @$"
DECLARE @TNBID INT = (SELECT TNB_TnbId FROM cdn.TypNieobec WHERE TNB_Nazwa = @NazwaNieobecnosci);
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
               (@PRI_PraId
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
    public static readonly string sqlQueryInsertPlanDoOptimy = $@"
DECLARE @id int;
DECLARE @EXISTSDZIEN INT = (SELECT COUNT([CDN].[PracPlanDni].[PPL_Data]) FROM cdn.PracPlanDni WHERE cdn.PracPlanDni.PPL_PraId = @PRI_PraId and [CDN].[PracPlanDni].[PPL_Data] = @DataInsert)
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
        ,ISNULL((SELECT TOP 1 KAD_TypDnia FROM cdn.KalendDni WHERE KAD_Data = @DataInsert), 1))
END TRY
BEGIN CATCH
END CATCH
END

SET @id = (select [cdn].[PracPlanDni].[PPL_PplId] from [cdn].[PracPlanDni] where [cdn].[PracPlanDni].[PPL_Data] = @DataInsert and [cdn].[PracPlanDni].[PPL_PraId] = @PRI_PraId);
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
	        @GodzOdDate,
	        @GodzDoDate,
	        4,
	        1,
	        1,
	        '');";
    public static readonly string sqlQueryInsertOdbNadgodzin = @"
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
		((select PPR_PprId from cdn.PracPracaDni where CAST(PPR_Data as datetime) = @DataInsert and PPR_PraId = @PRI_PraId),
		1,
		DATEADD(MINUTE, 0, @GodzOdDate),
		DATEADD(MINUTE, 0, @GodzDoDate),
		@TypPracy,
		1,
		1,
		'',
		@TypNadg);";
    public static readonly DateTime baseDate = new(1899, 12, 30);

    public static void Main()
    {
        while (true)
        {
            Check_Base_Files(); // sprawdz czy istnieje plik config, jesli nie to go inicjalizuje
            GetConfigFromFile();
            string[] folders;
            List<string> allfolders = [];
            foreach (string Big_Folder in Files_Folders)
            {
                try
                {
                    folders = Directory.GetDirectories(Big_Folder);
                    if (!folders.Any())
                    {
                        Console.WriteLine($"Nie znaleziono żadnych folderów w: {Big_Folder} {DateTime.Now}");
                    }
                    else
                    {
                        foreach (string folder in folders)
                        {
                            Check_Foldery_Processing(folder); // sprawdz czy istnieją odpowiednie podfoldery, jesli nie to je inicjalizuje
                            allfolders.Add(folder);
                        }
                    }
                }
                catch
                {
                    error_logger.Set_Error_File_Path(Path.Combine(AppDomain.CurrentDomain.BaseDirectory));
                    error_logger.New_Custom_Error($"Nie znaleziono folderu {Big_Folder} {DateTime.Now}");
                    Console.WriteLine($"Nie znaleziono folderu {Big_Folder} {DateTime.Now}");
                    continue;
                }
            }
            while (true)
            {
                foreach (string folder in allfolders)
                {
                    error_logger.Set_Error_File_Path(Path.Combine(folder, "Errors"));
                    error_logger.Current_Processed_Files_Folder = Path.Combine(folder, "Processed_Files");
                    error_logger.Current_Bad_Files_Folder = Path.Combine(folder, "Bad_Files");
                    try
                    {
                        Do_The_Thing(folder);
                    }
                    catch (Exception ex)
                    {
                        error_logger.Set_Error_File_Path(Path.Combine(AppDomain.CurrentDomain.BaseDirectory));
                        error_logger.New_Custom_Error($"{ex.Message} {DateTime.Now}");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine($"{ex.Message} {DateTime.Now}");
                        Console.ForegroundColor = ConsoleColor.White;
                        continue;
                    }
                }
                Thread.Sleep(3000);
                Console.Clear();
            }
        }

    }
    public static void Do_The_Thing(string Folder_Path)
    {
        string[] filePaths = Directory.GetFiles(Folder_Path);
        if (filePaths.Length == 0) {
            Console.WriteLine($"Nie znaleziono żadnych plików w folderze {Folder_Path} {DateTime.Now}");
            return;
        }
        foreach (string current_filePath in filePaths)
        {
            string filePath = current_filePath;
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Czytanie: {Path.GetFileNameWithoutExtension(filePath)} {DateTime.Now}");
            Console.ForegroundColor = ConsoleColor.White;
            if (Path.GetExtension(filePath) == ".xlsb")
            {
                error_logger.New_Custom_Error($"Błąd dla pliku {filePath}: Program nie obsługuje plików o rozszerzeniu xlsb. Proszę o plik z rozszerzeniem xlsx.");
                Console.WriteLine($"Błąd dla pliku {filePath}: Program nie obsługuje plików o rozszerzeniu xlsb. Proszę o plik z rozszerzeniem xlsx.");
                MoveFile(current_filePath);
                continue;
            }else if (!Is_File_Xlsx(filePath))
            {
                try
                {
                    ConvertToXlsx(filePath, Path.ChangeExtension(filePath, ".xlsx"));
                    MoveFile(current_filePath);
                    filePath = Path.ChangeExtension(filePath, ".xlsx");
                }
                catch
                {
                    MoveFile(current_filePath);
                    continue;
                }
            }

            error_logger.Nazwa_Pliku = filePath;
            (string Last_Mod_Osoba, DateTime Last_Mod_Time) = Get_File_Meta_Info(filePath);
            if (Last_Mod_Osoba == "Error") {
                error_logger.New_Custom_Error($"Błąd w czytaniu {filePath}, nie można wczytać metadanych {DateTime.Now}");
            }
            error_logger.Last_Mod_Osoba = Last_Mod_Osoba;
            error_logger.Last_Mod_Time = Last_Mod_Time;
            int ilosc_zakladek = 0;
            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                Usun_Ukryte_Karty(workbook);
                ilosc_zakladek = workbook.Worksheets.Count;
                for (int i = 1; i <= ilosc_zakladek; i++)
                {
                    error_logger.Nr_Zakladki = i;
                    IXLWorksheet zakladka = workbook.Worksheet(i);
                    error_logger.Nazwa_Zakladki = zakladka.Name;
                    int typ_pliku = Get_Typ_Zakladki(zakladka);
                    if (typ_pliku == 0)
                    {
                        Copy_Bad_Sheet_To_Files_Folder(filePath, error_logger.Nr_Zakladki);
                        error_logger.New_Custom_Error("Nie rozpoznano tego rodzaju zakladki: " + error_logger.Nazwa_Pliku + " nr zakladki: " + error_logger.Nr_Zakladki + " nazwa zakladki: " + error_logger.Nazwa_Zakladki + $" Porada: Sprawdź czy nagłówki są uzupełnione {DateTime.Now}");
                        Console.WriteLine("Nie rozpoznano tego rodzaju zakladki: " + error_logger.Nazwa_Pliku + " nr zakladki: " + error_logger.Nr_Zakladki + " nazwa zakladki: " + error_logger.Nazwa_Zakladki + $" Porada: Sprawdź czy nagłówki są uzupełnione {DateTime.Now}");
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
                    }else if (typ_pliku == 4)
                    {
                        try
                        {
                            Grafik_Pracy_Reader_2024_v2.Process_Zakladka_For_Optima(zakladka);
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
            string destinationPath = Path.Combine(error_logger.Current_Processed_Files_Folder, Path.GetFileName(filePath));

            if (File.Exists(destinationPath))
            {
                File.Delete(destinationPath);
            }

            File.Move(filePath, destinationPath);
        }
        catch
        {
        }
    }
    private static int Get_Typ_Zakladki(IXLWorksheet workshit)
    {
        foreach (IXLCell cell in workshit.CellsUsed()) // karta v2
        {
            try
            {
                if (cell.GetString().Contains("Dzień"))
                {
                    return 1;
                }
            }
            catch { }
        }

        string cellValue3_1 = workshit.Cell(3, 1).Value.ToString();
        if (cellValue3_1.Trim().Contains("GRAFIK PRACY")) // grafik v2
        {
            return 2;
        }

        cellValue3_1 = workshit.Cell(1, 1).Value.ToString();
        if (cellValue3_1.Trim().StartsWith("GRAFIK PRACY MIESIĄC")) // grafik v2024 v2
        {
            return 4;
        }

        string cellValue1_1 = workshit.Cell(1, 1).Value.ToString();
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
            using (XLWorkbook workbook = new XLWorkbook(File_Path))
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
        string newFilePath = System.IO.Path.Combine(error_logger.Current_Bad_Files_Folder, "DO_POPRAWY_" + System.IO.Path.GetFileName(filePath));
        try
        {
            using (XLWorkbook originalwb = new(filePath))
            {
                IXLWorksheet sheetToCopy = originalwb.Worksheet(sheetIndex);
                string newSheetName = sheetToCopy.Name;
                if (newSheetName.Length > 31)
                {
                    newSheetName = newSheetName.Substring(0, 31);
                }
                using (XLWorkbook workbook = File.Exists(newFilePath) ? new XLWorkbook(newFilePath) : new XLWorkbook())
                {
                    if (workbook.Worksheets.Contains(newSheetName))
                    {
                        return;
                    }
                    sheetToCopy.CopyTo(workbook, newSheetName);
                    XLWorkbookProperties properties = originalwb.Properties;
                    properties.Author = "Copied by program";
                    properties.Modified = DateTime.Now;
                    workbook.SaveAs(newFilePath);
                }
            }
        }
        catch
        {
        }
    }
    private static void Usun_Ukryte_Karty(XLWorkbook workbook)
    {
        List<IXLWorksheet> hiddenSheets = new List<IXLWorksheet>();
        foreach (IXLWorksheet sheet in workbook.Worksheets)
        {
            if (sheet.Visibility == XLWorksheetVisibility.Hidden)
            {
                hiddenSheets.Add(sheet);
            }
        }
        foreach (IXLWorksheet sheet in hiddenSheets)
        {
            workbook.Worksheets.Delete(sheet.Name);
        }
        workbook.Save();
    }
    private static void Check_Base_Files()
    {
        if (!File.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json")))
        {
            File.Create(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json")).Dispose();
            var defaultConfig = new
            {
                Files_Folders = new[] { Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Files") },
                Optima_Conection_String,
                Clear_Logs_On_Program_Restart,
                Clear_Processed_Files_On_Restart,
                Clear_Bad_Files_On_Restart
            };
            File.WriteAllText(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json"), JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
        }
    }
    private static void Check_Foldery_Processing(string FolderPath)
    {
        if (!Directory.Exists(Path.Combine(FolderPath, "Errors")))
        {
            try
            {
                Directory.CreateDirectory(Path.Combine(FolderPath, "Errors"));
                File.Create(Path.Combine(FolderPath, "Errors", "Errors.txt"));
            }
            catch
            {
                error_logger.New_Custom_Error($"Błędna ścieżka: {Path.Combine(FolderPath, "Errors")} {DateTime.Now}");
                Console.WriteLine($"Błędna ścieżka: {Path.Combine(FolderPath, "Errors")} {DateTime.Now}");
            }
        }

        if (File.Exists(Path.Combine(FolderPath, "Errors", "Errors.txt")) && Clear_Logs_On_Program_Restart)
        {
            File.WriteAllText(Path.Combine(FolderPath, "Errors", "Errors.txt"), string.Empty);
        }

        if (!Directory.Exists(Path.Combine(FolderPath, "Bad_Files")))
        {
            try
            {
                Directory.CreateDirectory(Path.Combine(FolderPath, "Bad_Files"));
            }
            catch
            {
                error_logger.New_Custom_Error($"Błędna ścieżka: {Path.Combine(FolderPath, "Bad_Files")} {DateTime.Now}");
                Console.WriteLine($"Błędna ścieżka: {Path.Combine(FolderPath, "Bad_Files")} {DateTime.Now}");
            }
        }
        else
        {
            if (Clear_Bad_Files_On_Restart)
            {
                foreach (string file in Directory.GetFiles(Path.Combine(FolderPath, "Bad_Files")))
                {
                    File.Delete(file);
                }
                foreach (string directory in Directory.GetDirectories(Path.Combine(FolderPath, "Bad_Files")))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }

        if (!Directory.Exists(Path.Combine(FolderPath, "Processed_Files")))
        {
            try
            {
                Directory.CreateDirectory(Path.Combine(FolderPath, "Processed_Files"));
            }
            catch
            {
                error_logger.New_Custom_Error($"Błędna ścieżka: {Path.Combine(FolderPath, "Processed_Files")} {DateTime.Now}");
                Console.WriteLine($"Błędna ścieżka: {Path.Combine(FolderPath, "Processed_Files")} {DateTime.Now}");
            }
        }
        else
        {
            if (Clear_Processed_Files_On_Restart)
            {
                foreach (string file in Directory.GetFiles(Path.Combine(FolderPath, "Processed_Files")))
                {
                    File.Delete(file);
                }
                foreach (string directory in Directory.GetDirectories(Path.Combine(FolderPath, "Processed_Files")))
                {
                    Directory.Delete(directory, recursive: true);
                }
            }
        }
    }
    private static bool Is_File_Xlsx(string filePath)
    {
        try
        {
            XLWorkbook workbook = new(filePath);
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
        // nwm dlaczego textwrap jest zawsze true. Jebać to jest wystarczająco dobre.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        DataSet dataSet;
        using (FileStream stream = File.Open(inputFilePath, FileMode.Open, FileAccess.Read))
        using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
        {
            ExcelDataSetConfiguration config = new()
            {
                ConfigureDataTable = _ => new ExcelDataTableConfiguration
                {
                    UseHeaderRow = true
                }
            };
            dataSet = reader.AsDataSet(config);
        }

        using XLWorkbook workbook = new XLWorkbook();
        foreach (System.Data.DataTable table in dataSet.Tables)
        {
            IXLWorksheet worksheet = workbook.Worksheets.Add(table.TableName);
            for (int i = 0; i < table.Columns.Count; i++)
                worksheet.Cell(1, i + 1).Value = table.Columns[i].ColumnName;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    object value = table.Rows[i][j];

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
        (string o, DateTime d) = Get_File_Meta_Info(inputFilePath);
        workbook.Properties.LastModifiedBy = o;
        workbook.Properties.Modified = d;
        workbook.SaveAs(outputFilePath);
    }
    public static void GetConfigFromFile()
    {
        string filePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Config.json");
        if (!File.Exists(filePath))
        {
            File.Create(filePath).Dispose();
            var defaultConfig = new
            {
                Optima_Conection_String,
                Clear_Processed_Files_On_Restart,
                Clear_Bad_Files_On_Restart,
                Clear_Logs_On_Program_Restart
            };
            File.WriteAllText(filePath, JsonSerializer.Serialize(defaultConfig, new JsonSerializerOptions { WriteIndented = true }));
        }
        string json = File.ReadAllText(filePath);
        Config_Data_Json? config = JsonSerializer.Deserialize<Config_Data_Json>(json);
        if (config != null)
        {
            Files_Folders = config.Files_Folders;
            Optima_Conection_String = config.Optima_Conection_String;
            Clear_Logs_On_Program_Restart = config.Clear_Logs_On_Program_Restart;
            Clear_Bad_Files_On_Restart = config.Clear_Bad_Files_On_Restart;
            Clear_Processed_Files_On_Restart = config.Clear_Processed_Files_On_Restart;
        }
    }
}
public class Config_Data_Json
{
    public List<string> Files_Folders { get; set; } = [];
    public string Optima_Conection_String { get; set; } = string.Empty;
    public bool Clear_Logs_On_Program_Restart { get; set;  } = false;
    public bool Clear_Bad_Files_On_Restart { get; set; } = true;
    public bool Clear_Processed_Files_On_Restart { get; set; } = true;
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