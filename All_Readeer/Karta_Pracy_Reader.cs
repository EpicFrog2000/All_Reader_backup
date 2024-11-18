using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;

namespace All_Readeer
{
    internal class Karta_Pracy_Reader
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
        }
        private class Karta_Pracy
        {
            public string Nazwa_Pliku = "";
            public int Nr_zakladki = 0;
            public Pracownik Pracownik { get; set; } = new();
            public string Oddzial { get; set; } = "";
            public string Zespol { get; set; } = "";
            public string Stanowisko { get; set; } = "";
            public int Nominal_Miesieczny_Ogolem { get; set; }
            public int Miesiac { get; set; } = 0;
            public int Rok { get; set; } = 0;
            public double Razem_Czas_Faktyczny_Przepracowany { get; set; } = 0;
            public double Razem_Praca_WG_Grafiku { get; set; } = 0;
            public double Razem_Przekr_Normy_Dobowej { get; set; } = 0;
            public double Razem_Ilosc_Godzin_Z_Dodatkiem_50 { get; set; } = 0;
            public double Razem_Ilosc_Godzin_Z_Dodatkiem_100 { get; set; } = 0;
            public double Razem_Godziny_W_Niedziele { get; set; } = 0;
            public double Razem_Godziny_W_Swieta { get; set; } = 0;
            public double Razem_Godziny_W_Nocy { get; set; } = 0;
            public double Razem_Dodatek_Szkodliwy_Ilosc_Godzin { get; set; } = 0;
            public double Praca_Po_Absencji { get; set; } = 0;
            public double Ogolem_Godziny_Nadliczbowe { get; set; } = 0;
            public List<Dane_Dni> Dane_Dni { get; set; } = new();
            public List<double> Zmniejszyc_Nominal_Miesięczny_O_Godziny_Usprawiedliwionej_Nieobecnosci { get; set; } = new();
            public double Brak_Do_Nominalu { get; set; } = 0;
            public double Spoznienia { get; set; } = 0;
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
            public List<Nieobecnosc> ListaNieobecnosci { get; set; } = [];
        }
        private class Dane_Dni
        {
            public int Dzien { get; set; } = 0;
            public TimeSpan Godzina_Rozpoczęcia_Pracy { get; set; } = TimeSpan.Zero;
            public string Absencja { get; set; } = "";
            public TimeSpan Godzina_Zakończenia_Pracy { get; set; } = TimeSpan.Zero;
            public double Czas_Faktyczny_Przepracowany { get; set; } = 0;
            public double Praca_WG_Grafiku { get; set; } = 0;
            public double Przekr_Normy_Dobowej { get; set; } = 0;
            public double Ilosc_Godzin_Z_Dodatkiem_50 { get; set; } = 0;
            public double Ilosc_Godzin_Z_Dodatkiem_100 { get; set; } = 0;
            public double Godziny_W_Niedziele { get; set; } = 0;
            public double Godziny_W_Swieta { get; set; } = 0;
            public double Godziny_W_Nocy { get; set; } = 0;
            public double Dodatek_Szkodliwy_Ilosc_Godzin { get; set; } = 0;
            public string Dodatek_Szkodliwy_Rodzaj_czynnosci { get; set; } = "";
        }
        private class Legenda
        {
            public int Id_Kodu { get; set; } = 0;
            public string Kod { get; set; } = "";
            public string Opis { get; set; } = "";
        }
        private class Pos
        {
            public int Row = 1;
            public int Col = 1;
        }
        private class Nieobecnosc
        {
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
            public Pracownik pracownik = new();
            public int rok = 0;
            public int miesiac = 0;
            public int dzien = 0;
            public RodzajNieobecnosci rodzaj_absencji = 0;
        }
        private enum RodzajNieobecnosci
        {
            DE,     // Delegacja
            DM,     // Dodatkowy urlop macierzyński
            DR,     // Urlop rodzicielski
            IK,     // Izolacja - Koronawirus
            NB,     // Badania lekarskie - okresowe
            NN,     // Nieobecność nieusprawiedliwiona
            NR,     // Badania lekarskie - z tyt. niepełnosprawności
            NU,     // Nieobecność usprawiedliwiona
            OD,     // Oddelegowanie do prac w ZZ
            OG,     // Odbiór godzin dyżuru
            ON,     // Odbiór nadgodzin
            OO,     // Odbiór pracy w niedziele
            OP,     // Urlop opiekuńczy (niepłatny)
            OS,     // Odbiór pracujących sobót
            PP,     // Poszukiwanie pracy
            PZ,     // Praca zdalna okazjonalna
            SW,     // Urlop/zwolnienie z tyt. siły wyższej
            SZ,     // Szkolenie
            SP,     // Zwolniony z obowiązku świadcz. pracy
            U9,     // Urlop rodzicielski 9 tygodni
            UA,     // Długotrwały urlop bezpłatny
            UB,     // Urlop bezpłatny
            UC,     // Urlop ojcowski
            UD,     // Na opiekę nad dzieckiem art.K.P.188
            UJ,     // Ćwiczenia wojskowe
            UK,     // Urlop dla krwiodawcy
            UL,     // Służba wojskowa
            ULawnika, // Praca ławnika w sądzie
            UM,     // Urlop macierzyński
            UN,     // Urlop z tyt. niepełnosprawności
            UO,     // Urlop okolicznościowy
            UP,     // Dodatkowy urlop osoby represjonowanej
            UR,     // Dodatkowe dni na turnus rehabilitacyjny
            US,     // Urlop szkoleniowy
            UV,     // Urlop weterana
            UW,     // Urlop wypoczynkowy
            UY,     // Urlop wychowawczy
            UZ,     // Urlop na żądanie
            WY,     // Wypoczynek skazanego
            ZC,     // Opieka nad członkiem rodziny (ZLA)
            ZD,     // Opieka nad dzieckiem (ZUS ZLA)
            ZK,     // Opieka nad dzieckiem Koronawirus
            ZL,     // Zwolnienie lekarskie (ZUS ZLA)
            ZN,     // Zwolnienie lekarskie niepłatne (ZLA)
            ZP,     // Kwarantanna sanepid
            ZR,     // Zwolnienie na rehabilitację (ZUS ZLA)
            ZS,     // Zwolnienie szpitalne (ZUS ZLA)
            ZY,     // Zwolnienie powypadkowe (ZUS ZLA)
            ZZ      // Zwolnienie lek. (ciąża) (ZUS ZLA)
        }
        private string Connection_String = "";
        private IXLWorksheet worksheet = null!;
        private List<Legenda> Lista_Legenda = [];
        private Pos Current_Pos = new(){ Row = 1, Col = 1 };
        private Karta_Pracy karta_Pracy = new();
        private List<Karta_Pracy> karty_Pracy = [];
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Last_Mod_Osoba = "";
        private int Try_Set_Num(string strnumer)
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
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheetO, string last_Mod_Osoba, DateTime last_Mod_Time)
        {

            try
            {
                worksheet = worksheetO;
                Last_Mod_Osoba = last_Mod_Osoba;
                Last_Mod_Time = last_Mod_Time;
                Init_Legenda();
                Current_Pos.Row = 1;
                while (true)
                {
                    try
                    {
                        karta_Pracy = new();
                        karta_Pracy.Nazwa_Pliku = Program.error_logger.Nazwa_Pliku;
                        karta_Pracy.Nr_zakladki = Program.error_logger.Nr_Zakladki;
                        Pos Shit_Start = Wykryj_Start_Karty();
                        if (Shit_Start.Row == -1)
                        {
                            break;
                        }
                        Wyczytaj_Naglowek(Shit_Start);
                        Wczytaj_Dane_Miesiaca(Shit_Start);
                        karty_Pracy.Add(karta_Pracy);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                        throw;
                    }
                    Current_Pos.Row++;
                }
                if (karty_Pracy.Count() > 0)
                {
                    foreach (var karta in karty_Pracy)
                    {
                        try
                        {
                            Dodaj_Dane_Do_Optimy(karta);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            throw;
                        }
                    }
                }
            }
            catch
            {
                throw;
            }

        }
        private void Wczytaj_Dane_Miesiaca(Pos Karta_Pos_Start)
        {
            Karta_Pos_Start.Row += 2;
            while (true)
            {
                Dane_Dni dzien = new();
                var strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col).GetFormattedString();

                if (string.IsNullOrEmpty(strnumer))
                {
                    break;
                }

                dzien.Dzien = Try_Set_Num(strnumer.Trim());
                if (dzien.Dzien == 0)
                {
                    if (DateTime.TryParse(strnumer, out DateTime Data))
                    {
                        dzien.Dzien = Data.Day;
                    }
                    else
                    {
                        Program.error_logger.New_Error(strnumer, "dzień", Karta_Pos_Start.Col, Karta_Pos_Start.Row, "Nieprawidłowy dzień");
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }
                try
                {
                    var danei = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 3).GetValue<string>();
                    if (!string.IsNullOrEmpty(danei))
                    {
                        Nieobecnosc nieobecnosc = new();
                        if (RodzajNieobecnosci.TryParse(danei.ToUpper(), out RodzajNieobecnosci Rnieobecnosc))
                        {
                            nieobecnosc.rodzaj_absencji = Rnieobecnosc;
                            nieobecnosc.pracownik = karta_Pracy.Pracownik;
                            nieobecnosc.rok = karta_Pracy.Rok;
                            nieobecnosc.miesiac = karta_Pracy.Miesiac;
                            nieobecnosc.dzien = dzien.Dzien;
                        }
                        else
                        {
                            Program.error_logger.New_Error(danei, "kod nieobecnosci", Karta_Pos_Start.Col + 3, Karta_Pos_Start.Row, "Nieprawidłowy kod nieobecności");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        karta_Pracy.ListaNieobecnosci.Add(nieobecnosc);
                        Karta_Pos_Start.Row++;
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }

                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 1).GetFormattedString().Trim();
                if (!string.IsNullOrEmpty(strnumer))
                {
                    try
                    {
                        dzien.Godzina_Rozpoczęcia_Pracy = Reader.Try_Get_Date(strnumer);
                    }
                    catch(Exception ex)
                    {
                        Program.error_logger.New_Error(strnumer, "Godzina_Rozpoczęcia_Pracy", Karta_Pos_Start.Col + 1, Karta_Pos_Start.Row, ex.Message);
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }

                dzien.Absencja = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 2).GetFormattedString().Trim();

                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 3).GetFormattedString().Trim();
                if (!string.IsNullOrEmpty(strnumer))
                {
                    try
                    {
                        dzien.Godzina_Zakończenia_Pracy = Reader.Try_Get_Date(strnumer);
                    }
                    catch (Exception ex)
                    {
                        Program.error_logger.New_Error(strnumer, "Godzina_Zakończenia_Pracy", Karta_Pos_Start.Col + 1, Karta_Pos_Start.Row, ex.Message);
                        throw new Exception(Program.error_logger.Get_Error_String());
                    }
                }

                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 4).GetFormattedString();
                dzien.Czas_Faktyczny_Przepracowany = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 5).GetFormattedString();
                dzien.Praca_WG_Grafiku = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 6).GetFormattedString();
                dzien.Przekr_Normy_Dobowej = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 7).GetFormattedString();
                dzien.Ilosc_Godzin_Z_Dodatkiem_50 = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 8).GetFormattedString();
                dzien.Ilosc_Godzin_Z_Dodatkiem_100 = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 9).GetFormattedString();
                dzien.Godziny_W_Niedziele = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 10).GetFormattedString();
                dzien.Godziny_W_Swieta = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 11).GetFormattedString();
                dzien.Godziny_W_Nocy = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 12).GetFormattedString();
                dzien.Dodatek_Szkodliwy_Ilosc_Godzin = Try_Set_Num(strnumer.Trim());
                dzien.Dodatek_Szkodliwy_Rodzaj_czynnosci = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 13).GetFormattedString().Trim();
                if (dzien.Godzina_Rozpoczęcia_Pracy!= TimeSpan.Zero && dzien.Godzina_Zakończenia_Pracy!= TimeSpan.Zero)
                {
                    karta_Pracy.Dane_Dni.Add(dzien);
                }
                Karta_Pos_Start.Row++;
            }
        }
        private void Wyczytaj_Naglowek(Pos Karta_Pos_Start)
        {
            Wczytaj_Pracownika(Karta_Pos_Start);
            if (karta_Pracy.Pracownik.Nazwisko == "" || karta_Pracy.Pracownik.Imie == "")
            {
                Program.error_logger.New_Error(karta_Pracy?.Pracownik?.Imie + karta_Pracy?.Pracownik?.Nazwisko, "Imie i Nazwisko", Karta_Pos_Start.Row, Karta_Pos_Start.Col);
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            var data = "";
            data = worksheet.Cell(Karta_Pos_Start.Row - 2, Karta_Pos_Start.Col + 12).GetFormattedString().Trim().Replace("   ", " ").Replace("  ", " ");
            if (data.EndsWith("r"))
            {
                data = data.Substring(0, data.Length - 1).Trim();
            }
            if (DateTime.TryParse(data, out DateTime parsedData))
            {
                karta_Pracy.Miesiac = parsedData.Month;
                karta_Pracy.Rok = parsedData.Year;
            }
            else
            {
                if (!string.IsNullOrEmpty(data) && data.Split(" ").Length > 1)
                {
                    var ndata = data.Split(" ");
                    if (ndata[0].ToLower() == "październik")
                    {
                        ndata[0] = "październik";
                    }
                    try
                    {
                        karta_Pracy.Set_Miesiac(ndata[0]);
                        karta_Pracy.Rok = Try_Set_Num(ndata[1]);
                    }
                    catch
                    {
                        karta_Pracy.Set_Miesiac("Zle dane");
                        karta_Pracy.Rok = Try_Set_Num(ndata[0]);
                    }
                }
                else
                {
                    data = worksheet.Cell(Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 12).GetFormattedString().Trim().Replace("   ", " ").Replace("  ", " ");
                    if (data.EndsWith("r"))
                    {
                        data = data.Substring(0, data.Length - 1).Trim();
                    }
                    if (!string.IsNullOrEmpty(data) && data.Split(" ").Length > 1)
                    {
                        var ndata = data.Split(" ");
                        try
                        {
                            karta_Pracy.Set_Miesiac(ndata[0]);
                            karta_Pracy.Rok = Try_Set_Num(ndata[1]);
                        }
                        catch
                        {
                            karta_Pracy.Set_Miesiac("Zle dane");
                            karta_Pracy.Rok = Try_Set_Num(ndata[0]);
                        }
                    }
                    else
                    {
                        if (DateTime.TryParse(data, out DateTime DataP))
                        {
                            karta_Pracy.Miesiac = DataP.Month;
                            karta_Pracy.Rok = DataP.Year;
                        }
                    }
                }
            }

            if (karta_Pracy.Miesiac == 0)
            {
                data = worksheet.Cell(Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 12).GetFormattedString().Trim().Replace("   ", " ").Replace("  ", " ");
                if (data.EndsWith("r"))
                {
                    data = data.Substring(0, data.Length - 1).Trim();
                }
                if (DateTime.TryParse(data, out DateTime parsedData2))
                {
                    karta_Pracy.Miesiac = parsedData2.Month;
                    karta_Pracy.Rok = parsedData2.Year;
                }
                else
                {
                    if (!string.IsNullOrEmpty(data) && data.Split(" ").Length > 1)
                    {
                        var ndata = data.Split(" ");
                        if (ndata[0].ToLower() == "październik")
                        {
                            ndata[0] = "październik";
                        }
                        try
                        {
                            karta_Pracy.Set_Miesiac(ndata[0]);
                            karta_Pracy.Rok = Try_Set_Num(ndata[1]);
                        }
                        catch
                        {
                            karta_Pracy.Set_Miesiac("Zle dane");
                            karta_Pracy.Rok = Try_Set_Num(ndata[0]);
                        }
                    }
                    else
                    {
                        data = worksheet.Cell(Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 12).GetFormattedString().Trim().Replace("   ", " ").Replace("  ", " ");
                        if (data.EndsWith("r"))
                        {
                            data = data.Substring(0, data.Length - 1).Trim();
                        }
                        if (!string.IsNullOrEmpty(data) && data.Split(" ").Length > 1)
                        {
                            var ndata = data.Split(" ");
                            try
                            {
                                karta_Pracy.Set_Miesiac(ndata[0]);
                                karta_Pracy.Rok = Try_Set_Num(ndata[1]);
                            }
                            catch
                            {
                                karta_Pracy.Set_Miesiac("Zle dane");
                                karta_Pracy.Rok = Try_Set_Num(ndata[0]);
                            }
                        }
                        else
                        {
                            if (DateTime.TryParse(data, out DateTime DataP))
                            {
                                karta_Pracy.Miesiac = DataP.Month;
                                karta_Pracy.Rok = DataP.Year;
                            }
                        }
                    }
                }
            }

            if (karta_Pracy.Miesiac == 0)
            {
                Program.error_logger.New_Error(data, "Miesiac", Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 12);
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            if (karta_Pracy.Rok == 0)
            {
                Program.error_logger.New_Error(data, "Rok", Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 12);
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private void Wczytaj_Pracownika(Pos Karta_Pos_Start)
        {
            karta_Pracy.Pracownik = new() { Imie = "", Nazwisko = "" };
            try
            {
                for (var i = 0; i < 10; i++)
                {
                    var imienazwisko = worksheet.Cell(Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 2 + i).GetFormattedString().Trim();
                    if (string.IsNullOrEmpty(imienazwisko))
                    {
                        continue;
                    }
                    else
                    {
                        var splited = imienazwisko.Split(" ");
                        Pracownik pracownik = new() { Imie = splited[1].Trim(), Nazwisko = splited[0].Trim() };
                        karta_Pracy.Pracownik = pracownik;
                        return;
                    }
                }
                var imienazwisko2 = worksheet.Cell(Karta_Pos_Start.Row - 2, Karta_Pos_Start.Col + 9).GetFormattedString().Trim();
                if (!string.IsNullOrEmpty(imienazwisko2))
                {
                    var splited = imienazwisko2.Split(" ");
                    Pracownik pracownik = new() { Imie = splited[1].Trim(), Nazwisko = splited[0].Trim() };
                    karta_Pracy.Pracownik = pracownik;
                    return;
                }
            }
            catch
            {
                return;
            }
        }
        private Pos Wykryj_Start_Karty()
        {
            int counter = 0;
            while (true)
            {
                if (counter > 100)
                {
                    return new() { Row = -1, Col = -1 };
                }
                try
                {
                    var cell = worksheet.Cell(Current_Pos.Row, Current_Pos.Col);
                    var cellval = "";
                    if (cell != null && !cell.IsEmpty())
                    {
                        cellval = cell.GetValue<string>().Trim();
                    }
                    else
                    {
                        cellval = string.Empty;
                    }
                    if (cellval == "Dz." || cellval == "Dzień")
                    {
                        return Current_Pos;
                    }
                }
                catch
                {
                    break;
                }
                Current_Pos.Row++;
                counter++;
            }

            return new() { Row = -1, Col = -1 };
        }
        private void Init_Legenda()
        {
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -1,
                Kod = "UW",
                Opis = "Urlop wypoczynkowy"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -2,
                Kod = "UB",
                Opis = "Urlop bezplatny"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -3,
                Kod = "UD",
                Opis = "2 dni opieki nad dzieckiem"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -4,
                Kod = "UZ",
                Opis = "Urlop na zadanie"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -5,
                Kod = "UO",
                Opis = "Urlop okolicznosciowy"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -6,
                Kod = "UM",
                Opis = "Urlop macierzynski"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -7,
                Kod = "CH",
                Opis = "Choroba"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -8,
                Kod = "ZC",
                Opis = "Opieka nad czlonkiem rodziny"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -9,
                Kod = "WZ",
                Opis = "Wezwanie do sadu, policji, innych organow"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -10,
                Kod = "ZD",
                Opis = "opieka nad dzieckiem do lat 14 - druk ZLA"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -11,
                Kod = "W",
                Opis = "Czas wolny za prace w godzinach nadliczbowych"
            });
            Lista_Legenda.Add(new Legenda
            {
                Id_Kodu = -12,
                Kod = "SP",
                Opis = "Zwolnienie z obowiazkow swiadczenia pracy"
            });
        }
        private void Insert_Obecnosci_do_Optimy(Karta_Pracy karta, SqlTransaction tran, SqlConnection connection)
        {
            if (!string.IsNullOrEmpty(karta.Pracownik.Imie) &&
                !string.IsNullOrEmpty(karta.Pracownik.Nazwisko) &&
                karta.Rok != 0 &&
                karta.Miesiac != 0)
            {
                NormalizeGodzinyZ2Dni(karta);
                foreach (var dzien in karta.Dane_Dni)
                {
                    try
                    {
                        using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{karta.Rok}-{karta.Miesiac:D2}-{dzien.Dzien:D2}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                            insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = "1899-12-30" + ' ' + dzien.Godzina_Rozpoczęcia_Pracy;
                            insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = "1899-12-30" + ' ' + dzien.Godzina_Zakończenia_Pracy;
                            insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", dzien.Czas_Faktyczny_Przepracowany);
                            insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", dzien.Praca_WG_Grafiku);
                            insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.Pracownik.Nazwisko);
                            insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.Pracownik.Imie);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_50", dzien.Ilosc_Godzin_Z_Dodatkiem_50);
                            insertCmd.Parameters.AddWithValue("@Godz_dod_100", dzien.Ilosc_Godzin_Z_Dodatkiem_100);
                            insertCmd.ExecuteScalar();
                        }
                    }
                    catch (SqlException ex)
                    {
                        tran.Rollback();
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    }
                    catch (FormatException)
                    {
                        tran.Rollback();
                        continue;
                    }
                    catch (Exception ex)
                    {
                        tran.Rollback();
                        if (ex.Data.Contains("kod") && ex.Data["kod"] is int kod && kod == 42069)
                        {
                            throw;
                        }
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    }
                }
            }
        }
        private void NormalizeGodzinyZ2Dni(Karta_Pracy karta)
        {
            for (int i = 0; i < karta.Dane_Dni.Count - 1; i++)
            {
                if (karta.Dane_Dni[i].Godzina_Rozpoczęcia_Pracy != TimeSpan.Zero &&
                    karta.Dane_Dni[i].Godzina_Zakończenia_Pracy == TimeSpan.Zero &&
                    karta.Dane_Dni[i + 1].Godzina_Rozpoczęcia_Pracy == TimeSpan.Zero &&
                    karta.Dane_Dni[i + 1].Godzina_Zakończenia_Pracy != TimeSpan.Zero)
                {
                    var Godzina_Polnocy_Nastepnego_Dnia = new TimeSpan(24, 0, 0);
                    var Godz_Pracy_Pierwszy_Dzien = Godzina_Polnocy_Nastepnego_Dnia - karta.Dane_Dni[i].Godzina_Rozpoczęcia_Pracy;
                    var Godz_Pracy = Godz_Pracy_Pierwszy_Dzien.Hours + (Godz_Pracy_Pierwszy_Dzien.Minutes / 60.0);

                    karta.Dane_Dni[i].Czas_Faktyczny_Przepracowany = Godz_Pracy;
                    karta.Dane_Dni[i].Praca_WG_Grafiku = Godz_Pracy;

                    karta.Dane_Dni[i + 1].Czas_Faktyczny_Przepracowany -= Godz_Pracy;
                    karta.Dane_Dni[i + 1].Praca_WG_Grafiku -= Godz_Pracy;
                }
            }
        }
        private int Ile_Dni_Roboczych(List<Nieobecnosc> listaNieobecnosci)
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
        private void Wjeb_Nieobecnosci_do_Optimy(List<Nieobecnosc> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
        {
            List<List<Nieobecnosc>> Nieobecnosci = Podziel_Niobecnosci_Na_Osobne(ListaNieobecności);
            foreach (var ListaNieo in Nieobecnosci)
            {
                var dni_robocze = Ile_Dni_Roboczych(ListaNieo);
                var dni_calosc = ListaNieo.Count;
                try
                {
                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertNieObecnoŚciDoOptimy, connection, tran))
                    {
                        DateTime dataBazowa = new DateTime(1899, 12, 30);
                        var nazwa_nieobecnosci = Dopasuj_Nieobecnosc(ListaNieo[0].rodzaj_absencji);
                        if (string.IsNullOrEmpty(nazwa_nieobecnosci))
                        {
                            Program.error_logger.New_Custom_Error($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                            var e = new Exception($"W programie brak dopasowanego kodu nieobecnosci: {ListaNieo[0].rodzaj_absencji} w dniu {new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien)} dla pracownika {ListaNieo[0].pracownik.Nazwisko} {ListaNieo[0].pracownik.Imie} z pliku: {Program.error_logger.Nazwa_Pliku} z zakladki: {Program.error_logger.Nr_Zakladki}. Nieobecnosc nie dodana.");
                            e.Data["Kod"] = 42069;
                            throw e;
                        }
                        DateTime dataniobecnoscistart = new DateTime(ListaNieo[0].rok, ListaNieo[0].miesiac, ListaNieo[0].dzien);
                        DateTime dataniobecnosciend = new DateTime(ListaNieo[ListaNieo.Count - 1].rok, ListaNieo[ListaNieo.Count - 1].miesiac, ListaNieo[ListaNieo.Count - 1].dzien);
                        int przyczyna = Dopasuj_Przyczyne(ListaNieo[0].rodzaj_absencji);
                        insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", ListaNieo[0].pracownik.Nazwisko);
                        insertCmd.Parameters.AddWithValue("@PracownikImieInsert", ListaNieo[0].pracownik.Imie);
                        insertCmd.Parameters.AddWithValue("@NazwaNieobecnosci", nazwa_nieobecnosci);
                        insertCmd.Parameters.AddWithValue("@DniPracy", dni_robocze);
                        insertCmd.Parameters.AddWithValue("@DniKalendarzowe", dni_calosc);
                        insertCmd.Parameters.AddWithValue("@Przyczyna", przyczyna);
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
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
                catch (FormatException)
                {
                    tran.Rollback();
                    continue;
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    if (ex.Data.Contains("kod") && ex.Data["kod"] is int kod && kod == 42069)
                    {
                        throw;
                    }
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
            }
        }
        private string Dopasuj_Nieobecnosc(RodzajNieobecnosci rodzaj)
        {
            return rodzaj switch
            {

                RodzajNieobecnosci.UO => "Urlop okolicznościowy",
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
                _ => "Nieobecność (B2B)"
            };
        }
        private int Dopasuj_Przyczyne(RodzajNieobecnosci rodzaj)
        {
            return rodzaj switch
            {
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
                _ => 9                             // Nie dotyczy dla pozostałych przypadków
            };
        }
        private List<List<Nieobecnosc>> Podziel_Niobecnosci_Na_Osobne(List<Nieobecnosc> listaNieobecnosci)
        {
            List<List<Nieobecnosc>> listaOsobnychNieobecnosci = new();
            List<Nieobecnosc> currentGroup = new();
            foreach (var nieobecnosc in listaNieobecnosci)
            {
                if (currentGroup.Count == 0 || nieobecnosc.dzien == currentGroup[^1].dzien + 1)
                {
                    currentGroup.Add(nieobecnosc);
                }
                else
                {
                    listaOsobnychNieobecnosci.Add(new List<Nieobecnosc>(currentGroup));
                    currentGroup = new List<Nieobecnosc> { nieobecnosc };
                }
            }
            if (currentGroup.Count > 0)
            {
                listaOsobnychNieobecnosci.Add(currentGroup);
            }
            return listaOsobnychNieobecnosci;
        }
        private void Dodaj_Dane_Do_Optimy(Karta_Pracy karta)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Connection_String))
                {
                    connection.Open();
                    SqlTransaction tran = connection.BeginTransaction();
                    Insert_Obecnosci_do_Optimy(karta, tran, connection);
                    Wjeb_Nieobecnosci_do_Optimy(karta.ListaNieobecnosci, tran, connection);
                    tran.Commit();
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno nieobecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.WriteLine($"Poprawnie dodawno obecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.ForegroundColor = ConsoleColor.White;
                    connection.Close();
                }
            }
            catch
            {
                throw;
            }
        }
    }
}
