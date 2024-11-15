using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace All_Readeer
{
    internal class Grafik_Pracy_Reader_v2
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
        }
        private class Grafik
        {
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
            public int rok { get; set; } = 0;
            public int miesiac { get; set; } = 0;
            public List<Dane_Dni> dane_dni = [];
            public List<Legenda> legenda = [];
            public void Set_Miesiac(string wartosc)
            {

                var mies = wartosc.Trim().ToLower();
                if (mies == "pażdziernik")
                {
                    mies = "październik";
                }
                if (mies == "styczeń")
                {
                    miesiac = 1;
                }else if (mies == "luty")
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
            public string kod = "";
        }
        private class Legenda {
            public int id = 0;
            public string kod = "";
            public string opis = "";
        }
        private class CurrentPosition
        {
            public int row { get; set; } = 1;
            public int col { get; set; } = 1;
        }
        private class Nieobecnosci
        {
            public int rok = 0;
            public int miesiac = 0;
            public int dzien = 0;
            public Pracownik pracownik = new();
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
        }
        private string File_Path = "";
        private string Last_Mod_Osoba = "";
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Optima_Connection_String = "";
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            if (string.IsNullOrEmpty(NewConnectionString))
            {
                Console.WriteLine("error: Empty Connection string");
                return;
            }
            Optima_Connection_String = NewConnectionString;
        }
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time)
        {
            try
            {
                Grafik grafik = new Grafik();
                grafik.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                grafik.nr_zakladki = Program.error_logger.Nr_Zakladki;
                Get_Header_Karta_Info(worksheet, ref grafik);
                Get_Dane_Dni(worksheet, ref grafik);
                Get_Legenda(worksheet, ref grafik);

                //get planowane nieobecnosci z grafiku
                List<Nieobecnosci> ListNieobecnosci = new();
                foreach (var dane_dni in grafik.dane_dni)
                {
                    foreach(var dzien in dane_dni.dane_dnia)
                    {
                        var matchingLegenda = grafik.legenda.FirstOrDefault(l => l.kod == dzien.kod);
                        if (matchingLegenda == null)
                        {
                            if (dzien.kod.Split(' ').Count() < 2)
                            {
                                Nieobecnosci nieobecnosci = new();
                                nieobecnosci.pracownik = dane_dni.pracownik;
                                nieobecnosci.rok = grafik.rok;
                                nieobecnosci.miesiac = grafik.miesiac;
                                nieobecnosci.dzien = dzien.dzien;
                                nieobecnosci.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                                nieobecnosci.nr_zakladki = Program.error_logger.Nr_Zakladki;
                                ListNieobecnosci.Add(nieobecnosci);
                            }
                        }
                    }
                }

                Dodaj_Dane_Do_Optimy(grafik, ListNieobecnosci);
            }
            catch
            {
                throw;
            }
        }
        private void Get_Header_Karta_Info(IXLWorksheet worksheet, ref Grafik grafik)
        {
            var dane = worksheet.Cell(3, 1).GetValue<string>().Trim();
            dane = Regex.Replace(dane, @"\s{2,}", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Tytułu Grafiku", 3, 1, "Brak Tytułu Grafiku");
                Console.WriteLine(Program.error_logger.Get_Error_String());
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            bool isParsed = int.TryParse(dane.Split(' ')[7], out int rok);
            if (!isParsed)
            {
                Program.error_logger.New_Error(dane, "Data Grafiku", 3, 1, "Błąd czytania daty");
                Console.WriteLine(Program.error_logger.Get_Error_String());
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            grafik.rok = rok;
            grafik.Set_Miesiac(dane.Split(' ')[6]);
            if(grafik.miesiac == 0)
            {
                Program.error_logger.New_Error(dane.Split(' ')[6], "Data Grafiku miesiac", 3, 1, "Błąd czytania miesiaca");
                Console.WriteLine(Program.error_logger.Get_Error_String());
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private void Get_Dane_Dni(IXLWorksheet worksheet, ref Grafik grafik)
        {
            CurrentPosition pozycja = new();
            pozycja.row = 6;
            while(true)
            {
                pozycja.col = 1;
                try
                {
                    var nazwiskoiimie = worksheet.Cell(pozycja.row, pozycja.col).GetValue<string>().Trim();
                    var NEXTnazwiskoiimie = worksheet.Cell(pozycja.row+1, pozycja.col).GetValue<string>().Trim();
                    if(string.IsNullOrEmpty(nazwiskoiimie) && string.IsNullOrEmpty(NEXTnazwiskoiimie)){
                        break;
                    }
                    if (!string.IsNullOrEmpty(nazwiskoiimie.Trim()) && nazwiskoiimie.Trim().Split(' ').Length < 3)
                    {
                        Dane_Dni dane_dni = new();
                        dane_dni.pracownik.Nazwisko = nazwiskoiimie.Split(' ')[0].Trim();
                        dane_dni.pracownik.Imie = nazwiskoiimie.Split(' ')[1].Trim();
                        pozycja.col = 3;
                        while (true)
                        {
                            Dane_Dnia dane_dnia = new();
                            var nrDnia = worksheet.Cell(5, pozycja.col).GetValue<string>().Trim();
                            if (!string.IsNullOrEmpty(nrDnia.Trim()))
                            {
                                if (int.TryParse(nrDnia, out int parsedDzien))
                                {
                                    dane_dnia.dzien = parsedDzien;
                                }
                                else if (DateTime.TryParse(nrDnia, out DateTime Data))
                                {
                                    dane_dnia.dzien = Data.Day;
                                }
                                else
                                {
                                    Program.error_logger.New_Error(nrDnia, "dzien", pozycja.col, 5, "Błędny nr dnia");
                                    Console.WriteLine(Program.error_logger.Get_Error_String());
                                    throw new Exception(Program.error_logger.Get_Error_String());
                                }
                                if(dane_dnia.dzien > 31 || dane_dnia.dzien == 0)
                                {
                                    break;
                                }
                            }
                            else
                            {
                                break;
                            }
                            var kodzik = worksheet.Cell(pozycja.row, pozycja.col).GetValue<string>().Trim();
                            if (!string.IsNullOrEmpty(kodzik))
                            {
                                dane_dnia.kod = kodzik;
                                if (dane_dnia.kod.Contains('.'))
                                {
                                    dane_dnia.kod = dane_dnia.kod.Split('.')[0].Trim();
                                }
                                if(dane_dnia.kod == null)
                                {
                                    Program.error_logger.New_Error(kodzik, "KodAktywnosciDnia", pozycja.col, pozycja.row, "Blędny kod aktywności dnia");
                                    Console.WriteLine(Program.error_logger.Get_Error_String());
                                    //throw new Exception(Program.error_logger.Get_Error_String());
                                }
                                else
                                {
                                    dane_dni.dane_dnia.Add(dane_dnia);
                                }
                            }
                            pozycja.col++;
                            if(pozycja.col >= 31)
                            {
                                break;
                            }
                        }
                        grafik.dane_dni.Add(dane_dni);
                    }
                    pozycja.row++;
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    break;
                }
            }
        }
        private void Get_Legenda(IXLWorksheet worksheet, ref Grafik grafik)
        {
            int idcounter = 0;
            CurrentPosition poz = new(){row = 21,col = 4};
            while (true)
            {
                try
                {
                    if(poz.row > 100)
                    {
                        break;
                    }
                    var dane = worksheet.Cell(poz.row, poz.col).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(dane))
                    {
                        idcounter++;
                        Legenda legenda = new();
                        legenda.kod = dane.Split('-')[0].Trim().Split(' ')[0].Trim();
                        if(legenda.kod == null)
                        {
                            legenda.kod = dane.Split('-')[0].Trim().Split('.')[0].Trim();
                        }
                        if(legenda.kod == null)
                        {
                            legenda.kod = dane.Split('-')[0].Trim();
                        }
                        if (legenda.kod.Contains('.'))
                        {
                            legenda.kod = legenda.kod.Split('.')[0].Trim();
                        }
                        if(legenda.kod == null)
                        {
                            Program.error_logger.New_Error(dane, "Linijka w legendie", poz.col, poz.row, "Program nie potrafi odczytać tej legendy");
                            Console.WriteLine(Program.error_logger.Get_Error_String());
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        legenda.id = idcounter;
                        legenda.opis = dane;
                        grafik.legenda.Add(legenda);
                    }
                    poz.row++;
                }catch(Exception ex){
                    throw new Exception(ex.Message);
                }
            }
        }
        private void Wpierdol_Plan_do_Optimy(Grafik grafik, SqlConnection connection, SqlTransaction tran)
        {
            var sqlQuery = $@"
DECLARE @id int;

IF((select DISTINCT COUNT(PRI_PraId) from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1) > 1)
BEGIN
	DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników o takim samym imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
	THROW 50001, @ErrorMessageC, 1;
END
DECLARE @PRI_PraId INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRI_PraId IS NULL
BEGIN
	SET @PRI_PraId = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRI_PraId IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
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
            foreach (var dane_Dni in grafik.dane_dni)
            {
                foreach (var dzien in dane_Dni.dane_dnia)
                {
                    try
                    {
                        using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
                        {
                            TimeSpan godz_rozp_pracy = TimeSpan.Zero;
                            TimeSpan godz_zak_pracy = TimeSpan.Zero;
                            // tutaj znajdz godz rozp i zak
                            var matchingLegenda = grafik.legenda.FirstOrDefault(l => l.kod == dzien.kod);
                            if (matchingLegenda == null)
                            {

                            }
                            else
                            {
                                if (matchingLegenda.opis.Contains("praca w godz.")) {
                                    var tmp = matchingLegenda.opis.Split("praca w godz.")[1];
                                    var r = tmp.Split('-')[0].Trim();
                                    var z = tmp.Split('-')[1].Trim();
                                    if (TimeSpan.TryParse(r.Replace('.', ':') + ":00", out TimeSpan rTime) && TimeSpan.TryParse(z.Replace('.', ':') + ":00", out TimeSpan zTime))
                                    {
                                        godz_rozp_pracy = rTime;
                                        godz_zak_pracy = zTime;
                                    }
                                    else
                                    {
                                        Program.error_logger.New_Custom_Error($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        Console.WriteLine($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        var e = new Exception($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                                        throw e;
                                    }
                                }
                                else if (matchingLegenda.opis.Contains("praca w godz"))
                                {
                                    var tmp = matchingLegenda.opis.Split("praca w godz")[1];
                                    var r = tmp.Split('-')[0].Trim();
                                    var z = tmp.Split('-')[1].Trim();
                                    if (TimeSpan.TryParse(r.Replace('.', ':') + ":00", out TimeSpan rTime) && TimeSpan.TryParse(z.Replace('.', ':') + ":00", out TimeSpan zTime))
                                    {
                                        godz_rozp_pracy = rTime;
                                        godz_zak_pracy = zTime;
                                    }
                                    else
                                    {
                                        Program.error_logger.New_Custom_Error($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        Console.WriteLine($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        var e = new Exception($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                                        throw e;
                                    }
                                }
                            }
                            insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{grafik.rok}-{grafik.miesiac:D2}-{dzien.dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                            insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = ("1899-12-30 " + godz_rozp_pracy.ToString());
                            insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = ("1899-12-30 " + godz_zak_pracy.ToString());
                            insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", (godz_zak_pracy - godz_rozp_pracy).TotalHours);
                            insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", (godz_zak_pracy - godz_rozp_pracy).TotalHours);
                            insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", dane_Dni.pracownik.Nazwisko);
                            insertCmd.Parameters.AddWithValue("@PracownikImieInsert", dane_Dni.pracownik.Imie);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_50", 0);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_100", 0);
                            insertCmd.ExecuteScalar();
                        }
                    }
                    catch (SqlException ex)
                    {
                        Program.error_logger.New_Custom_Error(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                        if (ex.Number == 50000)
                        {
                            Console.ForegroundColor = ConsoleColor.Yellow;
                        }
                        if (ex.Number == 50001)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                        }
                        Console.WriteLine(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                        Console.ForegroundColor = ConsoleColor.White;
                        tran.Rollback();
                        var e =  new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                        ex.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                        throw e;
                    }
                }
            }
            tran.Commit();
        }
        private void Wjeb_Nieobecnosci_do_Optimy(List<Nieobecnosci> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
        {
            var sqlQuery = @$"
IF((select DISTINCT COUNT(PRI_PraId) from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1) > 1)
BEGIN
	DECLARE @ErrorMessageC NVARCHAR(500) = 'Jest 2 pracowników o takim samym imieniu i nazwisku: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
	THROW 50001, @ErrorMessageC, 1;
END
DECLARE @PRACID INT = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikImieInsert and PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Typ = 1);
IF @PRACID IS NULL
BEGIN
	SET @PRACID = (select DISTINCT PRI_PraId from cdn.Pracidx WHERE PRI_Imie1 = @PracownikNazwiskoInsert  and PRI_Nazwisko = @PracownikImieInsert and PRI_Typ = 1);
	IF @PRACID IS NULL
	BEGIN
		DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie: ' +@PracownikImieInsert + ' ' +  @PracownikNazwiskoInsert;
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
            List<List<Nieobecnosci>> Nieobecnosci = Podziel_Niobecnosci_Na_Osobne(ListaNieobecności);
            foreach (var ListaNieo in Nieobecnosci)
            {
                var dni_robocze = Ile_Dni_Roboczych(ListaNieo);
                var dni_calosc = ListaNieo.Count;
                try
                {
                    using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
                    {
                        DateTime dataBazowa = new DateTime(1899, 12, 30);
                        DateTime dataniobecnoscistart = new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien);
                        DateTime dataniobecnosciend = new DateTime(ListaNieo[ListaNieo.Count - 1].rok, ListaNieo[ListaNieo.Count - 1].miesiac, ListaNieo[ListaNieo.Count - 1].dzien);
                        insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", ListaNieo[0].pracownik.Nazwisko);
                        insertCmd.Parameters.AddWithValue("@PracownikImieInsert", ListaNieo[0].pracownik.Imie);
                        insertCmd.Parameters.AddWithValue("@NazwaNieobecnosci", "Nieobecność (B2B) (plan)");
                        insertCmd.Parameters.AddWithValue("@DniPracy", dni_robocze);
                        insertCmd.Parameters.AddWithValue("@DniKalendarzowe", dni_calosc);
                        insertCmd.Parameters.AddWithValue("@Przyczyna", 9); // "Nieobecności inne"
                        insertCmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = dataniobecnoscistart;
                        insertCmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = dataBazowa;
                        insertCmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = dataniobecnosciend;
                        if (Last_Mod_Osoba.Length > 20)
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Last_Mod_Osoba.Substring(0, 20));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Last_Mod_Osoba);
                        }
                        if (Last_Mod_Osoba.Length > 50)
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Last_Mod_Osoba.Substring(0, 50));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Last_Mod_Osoba);
                        }
                        insertCmd.Parameters.AddWithValue("@DataMod", Last_Mod_Time);
                        insertCmd.ExecuteScalar();
                    }
                }
                catch (SqlException ex)
                {
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    if (ex.Number == 50000)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                    }
                    if (ex.Number == 50001)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                    }
                    Console.WriteLine(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    Console.ForegroundColor = ConsoleColor.White; tran.Rollback();
                    var e = new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                    throw e;
                }
            }

        }
        private void Dodaj_Dane_Do_Optimy(Grafik grafik, List<Nieobecnosci> ListaNieobecności)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Optima_Connection_String))
                {
                    connection.Open();
                    SqlTransaction tran = connection.BeginTransaction();
                    Wjeb_Nieobecnosci_do_Optimy(ListaNieobecności, tran, connection);
                    Wpierdol_Plan_do_Optimy(grafik, connection, tran);
                    tran.Commit();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno planowane nieobecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.WriteLine($"Poprawnie dodawno plan z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.ForegroundColor = ConsoleColor.White;
                }
            }
            catch
            {
                throw;
            }
        }
        private int Ile_Dni_Roboczych(List<Nieobecnosci> listaNieobecnosci)
        {
            int total = 0;
            foreach (var nieobecnosc in listaNieobecnosci)
            {
                DateTime absenceDate = new DateTime(nieobecnosc.rok, nieobecnosc.miesiac, nieobecnosc.dzien);
                if (absenceDate.DayOfWeek != DayOfWeek.Saturday && absenceDate.DayOfWeek != DayOfWeek.Sunday)
                {
                    total++;
                }
            }
            return total;
        }
        private List<List<Nieobecnosci>> Podziel_Niobecnosci_Na_Osobne(List<Nieobecnosci> listaNieobecnosci)
        {
            List<List<Nieobecnosci>> listaOsobnychNieobecnosci = new();
            List<Nieobecnosci> currentGroup = new();

            foreach (var nieobecnosc in listaNieobecnosci)
            {
                if (currentGroup.Count == 0 || nieobecnosc.dzien == currentGroup[^1].dzien + 1)
                {
                    currentGroup.Add(nieobecnosc);
                }
                else
                {
                    listaOsobnychNieobecnosci.Add(new List<Nieobecnosci>(currentGroup));
                    currentGroup = new List<Nieobecnosci> { nieobecnosc };
                }
            }

            if (currentGroup.Count > 0)
            {
                listaOsobnychNieobecnosci.Add(currentGroup);
            }

            return listaOsobnychNieobecnosci;
        }

    }
}