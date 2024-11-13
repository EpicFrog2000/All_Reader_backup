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
        private string FilePath = "";
        private string Connection_String = "Server=ITEGER-NT;Database=nowaBazaZadanie;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
        private string Optima_Connection_String = "Server=ITEGER-NT;Database=CDN_Wars_5;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
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
        private (string, DateTime) Get_File_Meta_Info()
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
        public void Set_Errors_File_Folder(string NewfilePath)
        {
            Program.error_logger.Set_Error_File_Path(NewfilePath);
        }
        public void Set_File_Path(string NewfilePath)
        {
            FilePath = NewfilePath;
        }
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }
        public void Set_Db_Tables_ConnectionString(string NewConnectionString)
        {
            Connection_String = NewConnectionString;
        }
        public void Process_For_Db()
        {
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            if (Last_Mod_Osoba=="Error") { throw new Exception("Error reading file"); }
            Init_Legenda();
            Read_Xlsx();
            Insert_Legenda_To_Db(Lista_Legenda);
            List<Pracownik> pracownicy = karty_Pracy.Select(k => k.Pracownik).Distinct().ToList();
            Insert_Pracownicy_To_Db(pracownicy);
            foreach (var karta in karty_Pracy)
            {
                int id = Insert_Karta_To_Db(karta);
                Insert_Dni_To_Db(id, karta.Dane_Dni);
            }
        }
        public void Process_For_Optima()
        {
            (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info();
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            Init_Legenda();
            Read_Xlsx();
            foreach (var karta in karty_Pracy)
            {
                Insert_Obecnosci_do_Optimy(karta);
            }
        }
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time)
        {
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
                    Wyczytaj_Footer(Shit_Start);
                    Wczytaj_Dane_Miesiaca(Shit_Start);
                    karty_Pracy.Add(karta_Pracy);
                }
                catch (Exception ex)
                {
                    Program.error_logger.New_Custom_Error(ex.Message);
                    throw new Exception(ex.Message);
                }
                Current_Pos.Row++;
            }
            foreach (var karta in karty_Pracy)
            {
                Insert_Obecnosci_do_Optimy(karta);
            }


        }
        private void Read_Xlsx()
        {
            try
            {
                using (var workbook = new XLWorkbook(FilePath))
                {
                    Program.error_logger.Nazwa_Pliku = FilePath;
                    foreach (var tworksheet in workbook.Worksheets)
                    {
                        Program.error_logger.Nr_Zakladki++;
                        Current_Pos.Row = 1;
                        worksheet = tworksheet;
                        while (true)
                        {
                            try
                            {
                                karta_Pracy = new();
                                karta_Pracy.Nazwa_Pliku = FilePath;
                                karta_Pracy.Nr_zakladki = Program.error_logger.Nr_Zakladki;
                                Pos Shit_Start = Wykryj_Start_Karty();
                                if (Shit_Start.Row == -1)
                                {
                                    break;
                                }
                                Wyczytaj_Naglowek(Shit_Start);
                                Wyczytaj_Footer(Shit_Start);
                                Wczytaj_Dane_Miesiaca(Shit_Start);
                                karty_Pracy.Add(karta_Pracy);
                            }
                            catch (Exception ex)
                            {
                                Program.error_logger.New_Custom_Error(ex.Message);
                                break;
                            }
                            Current_Pos.Row++;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Program.error_logger.New_Custom_Error(ex.Message);
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
                    Program.error_logger.New_Error(strnumer, "dzień", Karta_Pos_Start.Col, Karta_Pos_Start.Row);
                    return;
                }
                dzien.Dzien = Try_Set_Num(strnumer.Trim());
                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 1).GetFormattedString();
                strnumer = strnumer.Trim().Replace('.', ':');
                TimeSpan time;
                if (!string.IsNullOrEmpty(strnumer))
                {
                    if (!string.IsNullOrEmpty(strnumer))
                    {
                        if (TimeSpan.TryParse(strnumer.Trim(), out time))
                        {
                            dzien.Godzina_Rozpoczęcia_Pracy = time;
                        }
                        else
                        {
                            Program.error_logger.New_Error(strnumer, "Godzina_Rozpoczęcia_Pracy", Karta_Pos_Start.Col, Karta_Pos_Start.Row);
                            Console.WriteLine(Program.error_logger.Get_Error_String() + " (Zły foramt czasu)");
                        }
                    }
                }

                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 2).GetFormattedString();
                dzien.Absencja = strnumer.Trim();

                strnumer = worksheet.Cell(Karta_Pos_Start.Row, Karta_Pos_Start.Col + 3).GetFormattedString();
                strnumer = strnumer.Trim().Replace('.', ':');
                if (!string.IsNullOrEmpty(strnumer))
                {
                    if (!string.IsNullOrEmpty(strnumer.Trim()))
                    {
                        if (TimeSpan.TryParse(strnumer.Trim(), out time))
                        {
                            dzien.Godzina_Zakończenia_Pracy = time;
                        }
                        else
                        {
                            Program.error_logger.New_Error(strnumer, "Godzina_Zakończenia_Pracy", Karta_Pos_Start.Col, Karta_Pos_Start.Row);
                            Console.WriteLine(Program.error_logger.Get_Error_String() + " (Zły foramt czasu)");
                        }
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
                karta_Pracy.Dane_Dni.Add(dzien);
                Karta_Pos_Start.Row++;
            }
        }
        private void Wyczytaj_Footer(Pos Karta_Pos_Start)
        {
            var strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 4).GetFormattedString();
            karta_Pracy.Razem_Czas_Faktyczny_Przepracowany = Try_Set_Num(strnumer.Trim());

            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 5).GetFormattedString();
            karta_Pracy.Razem_Praca_WG_Grafiku = Try_Set_Num(strnumer.Trim());
            foreach (var leg in Lista_Legenda)
            {
                if (strnumer.Trim() == leg.Kod)
                {
                    karta_Pracy.Razem_Praca_WG_Grafiku = leg.Id_Kodu;
                    break;
                }
            }
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 34, Karta_Pos_Start.Col + 5).GetFormattedString();
            karta_Pracy.Praca_Po_Absencji = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 35, Karta_Pos_Start.Col + 5).GetFormattedString();
            karta_Pracy.Ogolem_Godziny_Nadliczbowe = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 6).GetFormattedString();
            karta_Pracy.Razem_Przekr_Normy_Dobowej = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 7).GetFormattedString();
            karta_Pracy.Razem_Ilosc_Godzin_Z_Dodatkiem_50 = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 8).GetFormattedString();
            karta_Pracy.Razem_Ilosc_Godzin_Z_Dodatkiem_100 = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 9).GetFormattedString();
            karta_Pracy.Razem_Godziny_W_Niedziele = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 10).GetFormattedString();
            karta_Pracy.Razem_Godziny_W_Swieta = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 11).GetFormattedString();
            karta_Pracy.Razem_Godziny_W_Nocy = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 33, Karta_Pos_Start.Col + 12).GetFormattedString();
            karta_Pracy.Razem_Dodatek_Szkodliwy_Ilosc_Godzin = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 37, Karta_Pos_Start.Col + 3).GetFormattedString();
            karta_Pracy.Brak_Do_Nominalu = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 39, Karta_Pos_Start.Col + 3).GetFormattedString();
            karta_Pracy.Spoznienia = Try_Set_Num(strnumer.Trim());
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 39, Karta_Pos_Start.Col + 8).GetFormattedString();
            if (!string.IsNullOrEmpty(strnumer.Trim()))
            {
                var newstrnumer = strnumer.Trim().Split(" ");
                karta_Pracy.Zmniejszyc_Nominal_Miesięczny_O_Godziny_Usprawiedliwionej_Nieobecnosci.Add(Try_Set_Num(newstrnumer[0]));
            }
            else
            {
                karta_Pracy.Zmniejszyc_Nominal_Miesięczny_O_Godziny_Usprawiedliwionej_Nieobecnosci.Add(0);
            }
            strnumer = worksheet.Cell(Karta_Pos_Start.Row + 38, Karta_Pos_Start.Col + 8).GetFormattedString();
            if (!string.IsNullOrEmpty(strnumer.Trim()))
            {
                var newstrnumer = strnumer.Trim().Split(" ");
                karta_Pracy.Zmniejszyc_Nominal_Miesięczny_O_Godziny_Usprawiedliwionej_Nieobecnosci.Add(Try_Set_Num(newstrnumer[0]));
            }
            else
            {
                karta_Pracy.Zmniejszyc_Nominal_Miesięczny_O_Godziny_Usprawiedliwionej_Nieobecnosci.Add(0);
            }
        }
        private void Wyczytaj_Naglowek(Pos Karta_Pos_Start)
        {
            var value = worksheet.Cell(Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 1).GetFormattedString().Trim();
            if (!string.IsNullOrEmpty(value))
            {
                karta_Pracy.Oddzial = value;
            }
            else
            {
                Program.error_logger.New_Error(value, "Oddzial", Karta_Pos_Start.Col + 1, Karta_Pos_Start.Row - 3);
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            value = worksheet.Cell(Karta_Pos_Start.Row - 2, Karta_Pos_Start.Col + 5).GetFormattedString().Trim();
            if (!string.IsNullOrEmpty(value))
            {
                karta_Pracy.Zespol = value;
            }
            else
            {
                Program.error_logger.New_Error(value, "Zespol", Karta_Pos_Start.Row - 2, Karta_Pos_Start.Col + 5);
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            value = worksheet.Cell(Karta_Pos_Start.Row - 1, Karta_Pos_Start.Col + 3).GetFormattedString().Trim();
            if (!string.IsNullOrEmpty(value))
            {
                karta_Pracy.Stanowisko = value;
            }
            else
            {
                Program.error_logger.New_Error(value, "Stanowisko", Karta_Pos_Start.Row - 1, Karta_Pos_Start.Col + 3);
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            Wczytaj_Pracownika(Karta_Pos_Start);

            if (karta_Pracy?.Pracownik?.Imie.Length == 0 || karta_Pracy?.Pracownik?.Nazwisko.Length == 0 || string.IsNullOrEmpty(karta_Pracy?.Pracownik?.Imie) || string.IsNullOrEmpty(karta_Pracy?.Pracownik?.Nazwisko))
            {
                Program.error_logger.New_Error(karta_Pracy?.Pracownik?.Imie + karta_Pracy?.Pracownik?.Nazwisko, "Imie i Nazwisko", Karta_Pos_Start.Row, Karta_Pos_Start.Col);
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            var data = worksheet.Cell(Karta_Pos_Start.Row - 2, Karta_Pos_Start.Col + 12).GetFormattedString().Trim().Replace("   ", " ").Replace("  ", " ");
            if (!string.IsNullOrEmpty(data))
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
                data = worksheet.Cell(Karta_Pos_Start.Row - 3, Karta_Pos_Start.Col + 12).GetFormattedString().Trim().Replace("   ", " ").Replace("  ", " ");
                if (!string.IsNullOrEmpty(data))
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

            var nominal = worksheet.Cell(Karta_Pos_Start.Row - 1, Karta_Pos_Start.Col + 13).GetFormattedString();
            if (!string.IsNullOrEmpty(nominal.Trim()))
            {
                var s = nominal.Trim().Split(" ");
                karta_Pracy.Nominal_Miesieczny_Ogolem = Try_Set_Num(s[0].Trim());
            }
            else
            {
                karta_Pracy.Nominal_Miesieczny_Ogolem = 0;
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
                    var cellval = worksheet.Cell(Current_Pos.Row, Current_Pos.Col).GetFormattedString().Trim();
                    if (cellval == "Dz.")
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
        private void Insert_Legenda_To_Db(List<Legenda> legendy)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    string insertQuery = "INSERT INTO Karty_Pracy_Kucharska_Legenda (Id_Kodu, Kod, Opis, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                            "VALUES (@Id_Kodu, @Kod, @Opis, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Os);";
                    foreach (var legenda in legendy)
                    {
                        using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@Id_Kodu", legenda.Id_Kodu);
                            insertCmd.Parameters.AddWithValue("@Kod", legenda.Kod);
                            insertCmd.Parameters.AddWithValue("@Opis", legenda.Opis);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);
                            insertCmd.ExecuteNonQuery();
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
                        string checkQuery = "SELECT COUNT(1) FROM Karta_Pracy_Kucharska_Pracownicy WHERE Imie = @Imie AND Nazwisko = @Nazwisko";
                        using (SqlCommand checkCmd = new SqlCommand(checkQuery, connection, tran))
                        {
                            checkCmd.Parameters.AddWithValue("@Imie", pracownik.Imie);
                            checkCmd.Parameters.AddWithValue("@Nazwisko", pracownik.Nazwisko);

                            int count = (int)checkCmd.ExecuteScalar();

                            if (count == 0)
                            {
                                string insertQuery = "INSERT INTO Karta_Pracy_Kucharska_Pracownicy (Imie, Nazwisko, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) VALUES (@Imie, @Nazwisko, @Ostatnia_Modyfikacja_Data , @Ostatnia_Modyfikacja_Os)";
                                using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                                {
                                    insertCmd.Parameters.AddWithValue("@Imie", pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@Nazwisko", pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                                    insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);
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
        private int Insert_Karta_To_Db(Karta_Pracy karta)
        {
            int insertedId = -1;
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    var Id_Pracownika = 0;
                    string selectQuery = "SELECT Id_Pracownika FROM Karta_Pracy_Kucharska_Pracownicy WHERE Imie = @Imie AND Nazwisko = @Nazwisko;";
                    using (SqlCommand selectCmd = new SqlCommand(selectQuery, connection, tran))
                    {
                        selectCmd.Parameters.Add("@Imie", SqlDbType.NVarChar).Value = karta.Pracownik.Imie;
                        selectCmd.Parameters.Add("@Nazwisko", SqlDbType.NVarChar).Value = karta.Pracownik.Nazwisko;
                        object result = selectCmd.ExecuteScalar();
                        Id_Pracownika = Convert.ToInt32(result);
                    }
                    string insertQuery = @"INSERT INTO Karty_Pracy_Kucharska (
                                    Id_Pracownika,
                                    Oddzial,
                                    Zespol,
                                    Stanowisko,
                                    Nominal_Miesieczny_Ogolem,
                                    Miesiac,
                                    Rok,
                                    Razem_Czas_Faktyczny_Przepracowany,
                                    Razem_Praca_WG_Grafiku,
                                    Razem_Przekr_Normy_Dobowej,
                                    Razem_Ilosc_Godzin_Z_Dodatkiem_50,
                                    Razem_Ilosc_Godzin_Z_Dodatkiem_100,
                                    Razem_Godziny_W_Niedziele,
                                    Razem_Godziny_W_Swieta,
                                    Razem_Godziny_W_Nocy,
                                    Razem_Dodatek_Szkodliwy_Ilosc_Godzin,
                                    Praca_Po_Absencji,
                                    Ogolem_Godziny_Nadliczbowe,
                                    Brak_Do_Nominalu,
                                    Spoznienia,
                                    Ostatnia_Modyfikacja_Data,
                                    Ostatnia_Modyfikacja_Osoba
                                )
                                VALUES(
                                    @Id_Pracownika,
                                    @Oddzial,
                                    @Zespol,
                                    @Stanowisko,
                                    @Nominal_Miesieczny_Ogolem,
                                    @Miesiac,
                                    @Rok,
                                    @Razem_Czas_Faktyczny_Przepracowany,
                                    @Razem_Praca_WG_Grafiku,
                                    @Razem_Przekr_Normy_Dobowej,
                                    @Razem_Ilosc_Godzin_Z_Dodatkiem_50,
                                    @Razem_Ilosc_Godzin_Z_Dodatkiem_100,
                                    @Razem_Godziny_W_Niedziele,
                                    @Razem_Godziny_W_Swieta,
                                    @Razem_Godziny_W_Nocy,
                                    @Razem_Dodatek_Szkodliwy_Ilosc_Godzin,
                                    @Praca_Po_Absencji,
                                    @Ogolem_Godziny_Nadliczbowe,
                                    @Brak_Do_Nominalu,
                                    @Spoznienia,
                                    @Ostatnia_Modyfikacja_Data,
                                    @Ostatnia_Modyfikacja_Os
                                ); SELECT SCOPE_IDENTITY();";

                    using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                    {
                        insertCmd.Parameters.Add("@Id_Pracownika", SqlDbType.Int).Value = Id_Pracownika;
                        insertCmd.Parameters.Add("@Oddzial", SqlDbType.NVarChar).Value = karta.Oddzial;
                        insertCmd.Parameters.Add("@Zespol", SqlDbType.NVarChar).Value = karta.Zespol;
                        insertCmd.Parameters.Add("@Stanowisko", SqlDbType.NVarChar).Value = karta.Stanowisko;
                        insertCmd.Parameters.Add("@Nominal_Miesieczny_Ogolem", SqlDbType.Int).Value = karta.Nominal_Miesieczny_Ogolem;
                        insertCmd.Parameters.Add("@Miesiac", SqlDbType.Int).Value = karta.Miesiac;
                        insertCmd.Parameters.Add("@Rok", SqlDbType.Int).Value = karta.Rok;
                        insertCmd.Parameters.Add("@Razem_Czas_Faktyczny_Przepracowany", SqlDbType.Int).Value = karta.Razem_Czas_Faktyczny_Przepracowany;
                        insertCmd.Parameters.Add("@Razem_Praca_WG_Grafiku", SqlDbType.Int).Value = karta.Razem_Praca_WG_Grafiku;
                        insertCmd.Parameters.Add("@Razem_Przekr_Normy_Dobowej", SqlDbType.Int).Value = karta.Razem_Przekr_Normy_Dobowej;
                        insertCmd.Parameters.Add("@Razem_Ilosc_Godzin_Z_Dodatkiem_50", SqlDbType.Int).Value = karta.Razem_Ilosc_Godzin_Z_Dodatkiem_50;
                        insertCmd.Parameters.Add("@Razem_Ilosc_Godzin_Z_Dodatkiem_100", SqlDbType.Int).Value = karta.Razem_Ilosc_Godzin_Z_Dodatkiem_100;
                        insertCmd.Parameters.Add("@Razem_Godziny_W_Niedziele", SqlDbType.Int).Value = karta.Razem_Godziny_W_Niedziele;
                        insertCmd.Parameters.Add("@Razem_Godziny_W_Swieta", SqlDbType.Int).Value = karta.Razem_Godziny_W_Swieta;
                        insertCmd.Parameters.Add("@Razem_Godziny_W_Nocy", SqlDbType.Int).Value = karta.Razem_Godziny_W_Nocy;
                        insertCmd.Parameters.Add("@Razem_Dodatek_Szkodliwy_Ilosc_Godzin", SqlDbType.Int).Value = karta.Razem_Dodatek_Szkodliwy_Ilosc_Godzin;
                        insertCmd.Parameters.Add("@Praca_Po_Absencji", SqlDbType.Int).Value = karta.Praca_Po_Absencji;
                        insertCmd.Parameters.Add("@Ogolem_Godziny_Nadliczbowe", SqlDbType.Int).Value = karta.Ogolem_Godziny_Nadliczbowe;
                        insertCmd.Parameters.Add("@Brak_Do_Nominalu", SqlDbType.Int).Value = karta.Brak_Do_Nominalu;
                        insertCmd.Parameters.Add("@Spoznienia", SqlDbType.Int).Value = karta.Spoznienia;
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                        insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);

                        insertedId = Convert.ToInt32(insertCmd.ExecuteScalar());
                        tran.Commit();
                    }
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Console.WriteLine(ex.ToString());
                }
            }
            return insertedId;
        }
        private void Insert_Dni_To_Db(int Id_Karty, List<Dane_Dni> dni)
        {
            using (SqlConnection connection = new SqlConnection(Connection_String))
            {
                connection.Open();
                SqlTransaction tran = connection.BeginTransaction();
                try
                {
                    string insertQuery = "INSERT INTO Karty_Pracy_Kucharska_Dane_Dni (Id_Karty, Dzien, Godzina_Rozpoczęcia_Pracy, Absencja, Godzina_Zakonczenia_Pracy, Czas_Faktyczny_Przepracowany, Praca_WG_Grafiku, Przekr_Normy_Dobowej, Ilosc_Godzin_Z_Dodatkiem_50, Ilosc_Godzin_Z_Dodatkiem_100, Godziny_W_Niedziele, Godziny_W_Swieta, Godziny_W_Nocy, Dodatek_Szkodliwy_Ilosc_Godzin, Dodatek_Szkodliwy_Rodzaj_czynnosci, Ostatnia_Modyfikacja_Data, Ostatnia_Modyfikacja_Osoba) " +
                                            "VALUES (@Id_Karty, @Dzien, @Godzina_Rozpoczecia, @Absencja, @Godzina_Zakonczenia_Pracy, @Czas_Faktyczny_Przepracowany, @Praca_WG_Grafiku, @Przekr_Normy_Dobowej, @Ilosc_Godzin_Z_Dodatkiem_50, @Ilosc_Godzin_Z_Dodatkiem_100, @Godziny_W_Niedziele, @Godziny_W_Swieta, @Godziny_W_Nocy, @Dodatek_Szkodliwy_Ilosc_Godzin, @Dodatek_Szkodliwy_Rodzaj_czynnosci, @Ostatnia_Modyfikacja_Data, @Ostatnia_Modyfikacja_Os);";

                    foreach (var dzień in dni)
                    {
                        using (SqlCommand insertCmd = new SqlCommand(insertQuery, connection, tran))
                        {
                            insertCmd.Parameters.AddWithValue("@Id_Karty", Id_Karty);
                            insertCmd.Parameters.AddWithValue("@Dzien", dzień.Dzien);
                            insertCmd.Parameters.AddWithValue("@Godzina_Rozpoczecia", dzień.Godzina_Rozpoczęcia_Pracy);
                            insertCmd.Parameters.AddWithValue("@Absencja", dzień.Absencja);
                            insertCmd.Parameters.AddWithValue("@Godzina_Zakonczenia_Pracy", dzień.Godzina_Zakończenia_Pracy);
                            insertCmd.Parameters.AddWithValue("@Czas_Faktyczny_Przepracowany", dzień.Czas_Faktyczny_Przepracowany);
                            insertCmd.Parameters.AddWithValue("@Praca_WG_Grafiku", dzień.Praca_WG_Grafiku);
                            insertCmd.Parameters.AddWithValue("@Przekr_Normy_Dobowej", dzień.Przekr_Normy_Dobowej);
                            insertCmd.Parameters.AddWithValue("@Ilosc_Godzin_Z_Dodatkiem_50", dzień.Ilosc_Godzin_Z_Dodatkiem_50);
                            insertCmd.Parameters.AddWithValue("@Ilosc_Godzin_Z_Dodatkiem_100", dzień.Ilosc_Godzin_Z_Dodatkiem_100);
                            insertCmd.Parameters.AddWithValue("@Godziny_W_Niedziele", dzień.Godziny_W_Niedziele);
                            insertCmd.Parameters.AddWithValue("@Godziny_W_Swieta", dzień.Godziny_W_Swieta);
                            insertCmd.Parameters.AddWithValue("@Godziny_W_Nocy", dzień.Godziny_W_Nocy);
                            insertCmd.Parameters.AddWithValue("@Dodatek_Szkodliwy_Ilosc_Godzin", dzień.Dodatek_Szkodliwy_Ilosc_Godzin);
                            insertCmd.Parameters.AddWithValue("@Dodatek_Szkodliwy_Rodzaj_czynnosci", dzień.Dodatek_Szkodliwy_Rodzaj_czynnosci);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Data", Last_Mod_Time);
                            insertCmd.Parameters.AddWithValue("@Ostatnia_Modyfikacja_Os", Last_Mod_Osoba);
                            insertCmd.ExecuteNonQuery();
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
        private void Insert_Obecnosci_do_Optimy(Karta_Pracy karta)
        {
            if (!string.IsNullOrEmpty(karta.Pracownik.Imie) &&
                !string.IsNullOrEmpty(karta.Pracownik.Nazwisko) &&
                karta.Rok != 0 &&
                karta.Miesiac != 0)
            {
                NormalizeGodzinyZ2Dni(karta);
                var sqlQuery = $@"
                DECLARE @id int;

                -- dodawaina pracownika do pracx i init pracpracdni
                DECLARE @PRI_PraId INT = (SELECT DISTINCT PRI_PraId FROM CDN.Pracidx where PRI_Nazwisko = @PracownikNazwiskoInsert and PRI_Imie1 = @PracownikImieInsert and PRI_Typ = 1)

                IF @PRI_PraId IS NULL
                BEGIN
                        DECLARE @ErrorMessage NVARCHAR(500) = 'Brak takiego pracownika w bazie: ' + @PracownikNazwiskoInsert + ' ' + @PracownikImieInsert;
                        THROW 50000, @ErrorMessage, 1;
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
                using (SqlConnection connection = new SqlConnection(Optima_Connection_String))
                {
                    connection.Open();
                    SqlTransaction tran = connection.BeginTransaction();
                    foreach (var dzien in karta.Dane_Dni)
                    {
                        try
                        {
                            using (SqlCommand insertCmd = new SqlCommand(sqlQuery, connection, tran))
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
                            Program.error_logger.New_Custom_Error(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                            Console.WriteLine(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                            tran.Rollback();
                            var e = new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                            e.Data["zakladka"] = Program.error_logger.Nr_Zakladki;
                            throw e;
                        }
                    }
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno obecnosci usera z pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                    Console.ForegroundColor = ConsoleColor.White;
                    tran.Commit();
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
    }
}
