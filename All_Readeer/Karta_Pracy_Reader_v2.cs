using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Data;
using System.Globalization;
using System.Text.RegularExpressions;

namespace All_Readeer
{
    internal class Karta_Pracy_Reader_v2
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
        }
        private class Karta_Pracy
        {
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
            public Pracownik pracownik { get; set; } = new();
            public int rok { get; set; } = 0;
            public int miesiac { get; set; } = 0;
            public int Set_Data(string Wartosc)
            {
                if (string.IsNullOrEmpty(Wartosc))
                {
                    return 1;
                }
                try
                {
                    DateTime data;
                    if (!DateTime.TryParse(Wartosc, out data))
                    {
                        return 1;
                    }
                    rok = data.Year;
                    miesiac = data.Month;
                } catch {
                    return 1;
                }
                return 0;
            }
            public List<Dane_Dnia> dane_dni { get; set; } = [];
            public List<Nieobecnosc> ListaNieobecnosci { get; set; } = [];
            public void Set_Miesiac(string nazwa)
            {
                if (!string.IsNullOrEmpty(nazwa))
                {
                    if (nazwa.ToLower() == "styczeń")
                    {
                        miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "i")
                    {
                        miesiac = 1;
                    }
                    else if (nazwa.ToLower() == "luty")
                    {
                        miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "ii")
                    {
                        miesiac = 2;
                    }
                    else if (nazwa.ToLower() == "marzec")
                    {
                        miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "iii")
                    {
                        miesiac = 3;
                    }
                    else if (nazwa.ToLower() == "kwiecień")
                    {
                        miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "iv")
                    {
                        miesiac = 4;
                    }
                    else if (nazwa.ToLower() == "maj")
                    {
                        miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "v")
                    {
                        miesiac = 5;
                    }
                    else if (nazwa.ToLower() == "czerwiec")
                    {
                        miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "vi")
                    {
                        miesiac = 6;
                    }
                    else if (nazwa.ToLower() == "lipiec")
                    {
                        miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "vii")
                    {
                        miesiac = 7;
                    }
                    else if (nazwa.ToLower() == "sierpień")
                    {
                        miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "viii")
                    {
                        miesiac = 8;
                    }
                    else if (nazwa.ToLower() == "wrzesień")
                    {
                        miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "ix")
                    {
                        miesiac = 9;
                    }
                    else if (nazwa.ToLower() == "październik")
                    {
                        miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "x")
                    {
                        miesiac = 10;
                    }
                    else if (nazwa.ToLower() == "listopad")
                    {
                        miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "xi")
                    {
                        miesiac = 11;
                    }
                    else if (nazwa.ToLower() == "grudzień")
                    {
                        miesiac = 12;
                    }
                    else if (nazwa.ToLower() == "xii")
                    {
                        miesiac = 12;
                    }
                    else
                    {
                        miesiac = 0;
                    }
                }
            }


        }
        private class Dane_Dnia
        {
            public int dzien { get; set; } = 0;
            public TimeSpan godz_rozp_pracy { get; set; } = TimeSpan.Zero;
            public TimeSpan godz_zakoncz_pracy { get; set; } = TimeSpan.Zero;
            public decimal praca_wg_grafiku { get; set; } = 0;
            public decimal liczba_godz_przepracowanych { get; set; } = 0;
            public decimal Godz_nadl_platne_z_dod_50 { get; set; } = 0;
            public decimal Godz_nadl_platne_z_dod_100 { get; set; } = 0;
        }
        private class CurrentPosition
        {
            public int row { get; set; } = 1;
            public int col { get; set; } = 1;
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
        private string Last_Mod_Osoba = "";
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Optima_Connection_String = "";
        private int Offset = 0;
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            Optima_Connection_String = NewConnectionString;
        }
        public void Process_Zakladka_For_Optima(IXLWorksheet worksheet, string last_Mod_Osoba, DateTime last_Mod_Time, int Typ_Zakladki)
        {
            Offset = Typ_Zakladki;
            try
            {
                Last_Mod_Osoba = last_Mod_Osoba;
                Last_Mod_Time = last_Mod_Time;
                List<Karta_Pracy> karty_pracy = [];
                CurrentPosition pozycja = new();
                Find_Karta(ref pozycja, worksheet);
                Karta_Pracy karta_pracy = new();
                karta_pracy.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                karta_pracy.nr_zakladki = Program.error_logger.Nr_Zakladki;
                Nieobecnosc nieobecnosc = new();
                nieobecnosc.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                nieobecnosc.nr_zakladki = Program.error_logger.Nr_Zakladki;
                Get_Header_Karta_Info(pozycja, worksheet, ref karta_pracy);

                Get_Dane_Dni(pozycja, worksheet, ref karta_pracy);
                karty_pracy.Add(karta_pracy);
                if(karty_pracy.Count > 0)
                {
                    foreach (var karta in karty_pracy)
                    {
                        try
                        {
                            Dodaj_Dane_Do_Optimy(karta);
                        }
                        catch
                        {
                            throw;
                        }
                    }
                }
            }catch(Exception ex){
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        private void Find_Karta(ref CurrentPosition pozycja, IXLWorksheet worksheet)
        {
            pozycja.col = 2 - Offset;
            bool found = false;
            try
            {
                foreach (var cell in worksheet.Column(pozycja.col).CellsUsed())
                {
                    if (cell.GetValue<string>().Equals("Dzień", StringComparison.OrdinalIgnoreCase))
                    {
                        pozycja.row = cell.Address.RowNumber;
                        found = true;
                        return;
                    }
                }
                if (!found)
                {
                    throw new Exception("Nie znaleziono słowa 'Dzień' w kolumnie.");
                }
            }
            catch
            {
                throw;
            }
        }
        private void Get_Header_Karta_Info(CurrentPosition StartKarty, IXLWorksheet worksheet, ref Karta_Pracy karta_pracy)
        {
            //wczytaj date
            var dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + 4).GetValue<string>().Trim().ToLower();
            for (int i = 0; i < 12; i++)
            {
                if (string.IsNullOrEmpty(dane))
                {
                    dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + 4 + i).GetValue<string>().Trim().ToLower();
                }
                else
                {
                    //here try to get data i rok
                    if (dane.EndsWith("r"))
                    {
                        dane = dane.Substring(0, dane.Length - 1).Trim();
                    }
                    if (dane.EndsWith("r."))
                    {
                        dane = dane.Substring(0, dane.Length - 2).Trim();
                    }

                    if (DateTime.TryParse(dane, out DateTime parsedData))
                    {
                        karta_pracy.miesiac = parsedData.Month;
                        karta_pracy.rok = parsedData.Year;
                    }
                    else
                    {
                        if (dane.Contains("pażdziernik"))
                        {
                            dane = dane.Replace("pażdziernik", "październik");
                        }
                        if (karta_pracy.Set_Data(dane) == 1)
                        {
                            if (dane.Split(" ").Length == 2)
                            {
                                var ndata = dane.Split(" ");
                                try
                                {
                                    karta_pracy.Set_Miesiac(ndata[0]);
                                    if (int.TryParse(Regex.Replace(ndata[1], @"\D", ""), out int rok))
                                    {
                                        karta_pracy.rok = rok;
                                    }
                                }
                                catch{}
                            }
                            else if (dane.Split(" ").Length == 3)
                            {
                                var ndata = dane.Split(" ");
                                try
                                {
                                    karta_pracy.Set_Miesiac(ndata[1]);
                                    if (int.TryParse(ndata[2], out int rok))
                                    {
                                        karta_pracy.rok = rok;
                                    }
                                }
                                catch{}
                            }
                            else
                            {
                                if(dane.Split(" ").Count() > 1)
                                {
                                    //wez 2 od tylu
                                    var ndata = dane.Split(" ");
                                    try
                                    {
                                        karta_pracy.Set_Miesiac(ndata[^2]);
                                        if (int.TryParse(ndata[^1], out int rok))
                                        {
                                            karta_pracy.rok = rok;
                                        }
                                    }
                                    catch{}
                                }
                            }
                        }
                    }
                    if (karta_pracy.miesiac == 0 || karta_pracy.rok == 0)
                    {
                        dane = worksheet.Cell(StartKarty.row - 4, StartKarty.col + 4 + i - 1).GetValue<string>().Trim().ToLower();
                        if (!string.IsNullOrEmpty(dane) && DateTime.TryParse(dane, out DateTime parsedData2))
                        {
                            karta_pracy.miesiac = parsedData2.Month;
                            karta_pracy.rok = parsedData2.Year;
                        }

                    }

                    if (karta_pracy.miesiac != 0 && karta_pracy.rok != 0)
                    {
                        break;
                    }
                }
            }
            if (karta_pracy.miesiac == 0 || karta_pracy.rok == 0)
            {
                Program.error_logger.New_Error(dane, "data", StartKarty.col + 11, StartKarty.row - 3, "Nie wykryto daty w pliku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }


            //wczytaj nazwisko i imie
            string[] wordsToRemove = { "IMIĘ:", "IMIE:", "NAZWISKO:", "NAZWISKO", " IMIE", "IMIĘ" };
            dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col).GetValue<string>().Trim().Replace("  ", " ");
            for (int i = 0; i < 6; i++)
            {
                foreach (var word in wordsToRemove)
                {
                    var pattern = $@"\b{Regex.Escape(word)}\b";
                    dane = Regex.Replace(dane, pattern, "", RegexOptions.IgnoreCase);
                }

                dane = Regex.Replace(dane, @"\s+", " ").Trim();
                if (dane.Contains("KARTA PRACY:"))
                {
                    dane = dane.Replace("KARTA PRACY:", "").Trim();
                }
                if (!string.IsNullOrEmpty(dane))
                {
                    break;
                }
                else
                {
                    dane = worksheet.Cell(StartKarty.row - 2, StartKarty.col + i).GetValue<string>().Trim().Replace("  ", " ");
                }
            }
            if (string.IsNullOrEmpty(dane))
            {
                dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col).GetValue<string>().Trim().Replace("  ", " ");
                for (int i = 0; i < 6; i++)
                {
                    foreach (var word in wordsToRemove)
                    {
                        dane = dane.Replace(word, "", StringComparison.OrdinalIgnoreCase);
                    }
                    dane = Regex.Replace(dane, @"\s+", " ").Trim();
                    if (dane.Contains("KARTA PRACY:"))
                    {
                        dane = dane.Replace("KARTA PRACY:", "").Trim();
                    }
                    if (!string.IsNullOrEmpty(dane))
                    {
                        break;
                    }
                    else
                    {
                        dane = worksheet.Cell(StartKarty.row - 3, StartKarty.col + i).GetValue<string>().Trim().Replace("  ", " ");
                    }
                }
            }


            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Nie wykryto nazwiska i imienia w pliku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            foreach (var word in wordsToRemove)
            {
                dane = dane.Replace(word, "", StringComparison.OrdinalIgnoreCase);
            }
            dane = Regex.Replace(dane, @"\s+", " ").Trim();
            if (dane.Contains("KARTA PRACY:"))
            {
                dane = dane.Replace("KARTA PRACY:", "").Trim();
            }
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row -2, "Zły format pola nazwisko i imie");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            else
            {
                try
                {
                    karta_pracy.pracownik.Nazwisko = dane.Trim().Split(' ')[0];
                    karta_pracy.pracownik.Imie = dane.Trim().Split(' ')[1];
                }
                catch
                {
                    Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row - 2, "Zły format pola nazwisko i imie");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
            }
            if (karta_pracy.pracownik.Nazwisko == null || karta_pracy.pracownik.Imie == null)
            {
                Program.error_logger.New_Error(dane, "nazwisko i imie", StartKarty.col, StartKarty.row -2, "Zły format pola nazwisko i imie");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private void Get_Dane_Dni(CurrentPosition StartKarty, IXLWorksheet worksheet, ref Karta_Pracy karta_pracy)
        {
            StartKarty.row += 3;
            var NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
            while (!string.IsNullOrEmpty(NrDnia))
            {
                Dane_Dnia dzien = new();
                // dzien miesiaca
                if (int.TryParse(NrDnia, out int parsedDzien))
                {
                    dzien.dzien = parsedDzien;
                }else if (DateTime.TryParse(NrDnia, out DateTime Data))
                {
                    dzien.dzien = Data.Day;
                }else
                {
                    Program.error_logger.New_Error(NrDnia, "dzien", StartKarty.col, StartKarty.row, "Błędny nr dnia");
                    throw new Exception(Program.error_logger.Get_Error_String());
                }
                var Cell_Value = "";
                //try get nieobecność:
                try
                {
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 3).GetValue<string>();
                    if (!string.IsNullOrEmpty(Cell_Value.Trim()))
                    {
                        Nieobecnosc nieobecnosc = new();
                        if (RodzajNieobecnosci.TryParse(Cell_Value.ToUpper(), out RodzajNieobecnosci Rnieobecnosc))
                        {
                            nieobecnosc.rodzaj_absencji = Rnieobecnosc;
                            nieobecnosc.pracownik = karta_pracy.pracownik;
                            nieobecnosc.rok = karta_pracy.rok;
                            nieobecnosc.miesiac = karta_pracy.miesiac;
                            nieobecnosc.dzien = dzien.dzien;
                        }
                        else
                        {
                            Program.error_logger.New_Error(Cell_Value, "Kod absencji", StartKarty.col + 3, StartKarty.row, "Nieprawidłowy kod nieobecności");
                            throw new Exception(Program.error_logger.Get_Error_String());
                        }
                        karta_pracy.ListaNieobecnosci.Add(nieobecnosc);
                        StartKarty.row++;
                        NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
                        continue;
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception(ex.Message);
                }
                // godz rozpoczecia
                try{
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 1).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(Cell_Value))
                    {
                        dzien.godz_rozp_pracy = Reader.Try_Get_Date(Cell_Value);
                    }
                }
                catch(Exception ex)
                {
                    Program.error_logger.New_Error(Cell_Value, "Godzina_Rozpoczęcia_Pracy", StartKarty.col + 1, StartKarty.row, ex.Message);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                // godz zakonczenia
                try
                {
                    Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 2).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(Cell_Value))
                    {
                        dzien.godz_zakoncz_pracy = Reader.Try_Get_Date(Cell_Value);
                    }
                }
                catch(Exception ex)
                {
                    Program.error_logger.New_Error(Cell_Value, "Godzina_Rozpoczęcia_Pracy", StartKarty.col + 1, StartKarty.row, ex.Message);
                    throw new Exception(Program.error_logger.Get_Error_String());
                }

                //get godz_nad 50
                Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 9).GetValue<string>().Trim();
                if (!string.IsNullOrEmpty(Cell_Value))
                {
                    dzien.Godz_nadl_platne_z_dod_50 = decimal.Parse(Cell_Value);
                }
                //get godz_nad 100
                Cell_Value = worksheet.Cell(StartKarty.row, StartKarty.col + 10).GetValue<string>().Trim();
                if (!string.IsNullOrEmpty(Cell_Value))
                {
                    dzien.Godz_nadl_platne_z_dod_100 = decimal.Parse(Cell_Value);
                }


                if (dzien.godz_rozp_pracy != TimeSpan.Zero && dzien.godz_zakoncz_pracy != TimeSpan.Zero)
                {
                    karta_pracy.dane_dni.Add(dzien);
                }
                StartKarty.row++;
                NrDnia = worksheet.Cell(StartKarty.row, StartKarty.col).GetValue<string>().Trim();
            }
        }
        private void Wpierdol_Obecnosci_do_Optimy(Karta_Pracy karta, SqlTransaction tran, SqlConnection connection)
        {

            foreach (var dane_Dni in karta.dane_dni)
            {
                try
                {
                    DateTime WażnaData = DateTime.Parse($"{karta.rok}-{karta.miesiac:D2}-{dane_Dni.dzien:D2}");
                    var (startPodstawowy, endPodstawowy, startNadl50, endNadl50, startNadl100, endNadl100) = Oblicz_Czas_Z_Dodatkiem(dane_Dni);
                    double czasPrzepracowany = 0;
                    if (dane_Dni.godz_zakoncz_pracy < dane_Dni.godz_rozp_pracy)
                    {
                        czasPrzepracowany = (TimeSpan.FromHours(24) - dane_Dni.godz_rozp_pracy).TotalHours + dane_Dni.godz_zakoncz_pracy.TotalHours;
                    }
                    else
                    {
                        czasPrzepracowany = (dane_Dni.godz_zakoncz_pracy - dane_Dni.godz_rozp_pracy).TotalHours;
                    }
                    double czasPodstawowy = czasPrzepracowany - (double)(dane_Dni.Godz_nadl_platne_z_dod_50 + dane_Dni.Godz_nadl_platne_z_dod_100);

                    bool czy_next_dzien = false;

                    // zrob to co ponizej ale dla wszystkich 3 xdd
                    if (czasPodstawowy > 0)
                    {
                        if (endPodstawowy < startPodstawowy)
                        {
                            czy_next_dzien = true;
                            // insert godziny przed północą
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@DataInsert",WażnaData);
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = baseDate + startPodstawowy;
                                DateTime godzDoDate = baseDate + TimeSpan.FromHours(24);
                                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                double czasPrzepracowanyInsert = (TimeSpan.FromHours(24) - startPodstawowy).TotalHours;
                                insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                insertCmd.Parameters.AddWithValue("@TypPracy", 2); // podstawowy
                                insertCmd.ExecuteScalar();
                            }
                            // insert po północy
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                var data = WażnaData.AddDays(1);
                                insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{data:yyyy-MM-dd}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = baseDate + TimeSpan.FromHours(0);
                                DateTime godzDoDate = baseDate + endPodstawowy;
                                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                double czasPrzepracowanyInsert = endPodstawowy.TotalHours;
                                insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                insertCmd.Parameters.AddWithValue("@TypPracy", 2); // podstawowy
                                insertCmd.ExecuteScalar();
                            }
                        }
                        else
                        {
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@DataInsert", WażnaData);
                                DateTime dataBazowa = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = dataBazowa + startPodstawowy;
                                DateTime godzDoDate = dataBazowa + endPodstawowy;
                                insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", (endPodstawowy - startPodstawowy).TotalHours);
                                insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", (endPodstawowy - startPodstawowy).TotalHours);
                                insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                insertCmd.Parameters.AddWithValue("@TypPracy", 2); // podstawowy
                                insertCmd.ExecuteScalar();
                            }
                        }
                    }
                    if (czy_next_dzien)
                    {
                        WażnaData = WażnaData.AddDays(1);
                        czy_next_dzien = false;
                    }
                    if (dane_Dni.Godz_nadl_platne_z_dod_50 > 0)
                    {
                        if (endNadl50 < startNadl50)
                        {
                            czy_next_dzien = true;
                            // insert godziny przed północą
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = baseDate + startNadl50;
                                DateTime godzDoDate = baseDate + TimeSpan.FromHours(24);
                                double czasPrzepracowanyInsert = (TimeSpan.FromHours(24) - startNadl50).TotalHours;
                                if (godzOdDate != godzDoDate)
                                {
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@DataInsert", WażnaData);
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 8); // 50%
                                    insertCmd.ExecuteScalar();
                                }
                            }
                            // insert po północy
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                var data = WażnaData.AddDays(1);
                                insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{data:yyyy-MM-dd}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = baseDate + TimeSpan.FromHours(0);
                                DateTime godzDoDate = baseDate + endNadl50;
                                double czasPrzepracowanyInsert = endNadl50.TotalHours;
                                if (godzOdDate != godzDoDate)
                                {
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 8); // 50%
                                    insertCmd.ExecuteScalar();
                                }
                            }
                        }
                        else
                        {
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@DataInsert", WażnaData);

                                DateTime dataBazowa = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = dataBazowa + startNadl50;
                                DateTime godzDoDate = dataBazowa + endNadl50;
                                if (godzOdDate != godzDoDate)
                                {
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", (endNadl50 - startNadl50).TotalHours);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", (endNadl50 - startNadl50).TotalHours);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 8); // 50%
                                    insertCmd.ExecuteScalar();
                                }
                            }
                        }
                    }
                    if (czy_next_dzien)
                    {
                        WażnaData = WażnaData.AddDays(1);
                        czy_next_dzien = false;
                    }
                    if (dane_Dni.Godz_nadl_platne_z_dod_100 > 0)
                    {
                        czy_next_dzien = true;
                        if (endNadl100 < startNadl100)
                        {
                            // insert godziny przed północą
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@DataInsert", WażnaData);
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = baseDate + startNadl100;
                                DateTime godzDoDate = baseDate + TimeSpan.FromHours(24);
                                double czasPrzepracowanyInsert = (TimeSpan.FromHours(24) - startNadl100).TotalHours;
                                if (godzOdDate != godzDoDate)
                                {
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 6); // 100%
                                    insertCmd.ExecuteScalar();
                                }
                            }
                            // insert po północy
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                var data = WażnaData.AddDays(1);
                                insertCmd.Parameters.AddWithValue("@DataInsert", DateTime.ParseExact($"{data:yyyy-MM-dd}", "yyyy-MM-dd", CultureInfo.InvariantCulture));
                                DateTime baseDate = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = baseDate + TimeSpan.FromHours(0);
                                DateTime godzDoDate = baseDate + endNadl100;
                                double czasPrzepracowanyInsert = endNadl100.TotalHours;
                                if (godzOdDate != godzDoDate)
                                {
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", czasPrzepracowanyInsert);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 6); // 100%
                                    insertCmd.ExecuteScalar();
                                }
                            }
                        }
                        else
                        {
                            using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertObecnościDoOptimy, connection, tran))
                            {
                                insertCmd.Parameters.AddWithValue("@DataInsert", WażnaData);

                                DateTime dataBazowa = new DateTime(1899, 12, 30);
                                DateTime godzOdDate = dataBazowa + startNadl100;
                                DateTime godzDoDate = dataBazowa + endNadl100;
                                if (godzOdDate != godzDoDate)
                                {
                                    insertCmd.Parameters.Add("@GodzOdDate", SqlDbType.DateTime).Value = godzOdDate;
                                    insertCmd.Parameters.Add("@GodzDoDate", SqlDbType.DateTime).Value = godzDoDate;
                                    insertCmd.Parameters.AddWithValue("@CzasPrzepracowanyInsert", (endNadl100 - startNadl100).TotalHours);
                                    insertCmd.Parameters.AddWithValue("@PracaWgGrafikuInsert", (endNadl100 - startNadl100).TotalHours);
                                    insertCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", karta.pracownik.Nazwisko);
                                    insertCmd.Parameters.AddWithValue("@PracownikImieInsert", karta.pracownik.Imie);
                                    insertCmd.Parameters.AddWithValue("@TypPracy", 6); // 100%
                                    insertCmd.ExecuteScalar();
                                }
                            }
                        }
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
                    continue;
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}");
                }
            }
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
                        DateTime dataniobecnosciend = new DateTime(ListaNieo[ListaNieo.Count-1].rok, ListaNieo[ListaNieo.Count-1].miesiac, ListaNieo[ListaNieo.Count-1].dzien);
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
                using (SqlConnection connection = new SqlConnection(Optima_Connection_String))
                {
                    connection.Open();
                    SqlTransaction tran = connection.BeginTransaction();
                    Wpierdol_Obecnosci_do_Optimy(karta, tran, connection);
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
        private (TimeSpan, TimeSpan, TimeSpan, TimeSpan, TimeSpan, TimeSpan) Oblicz_Czas_Z_Dodatkiem(Dane_Dnia dane_Dni)
        {
            TimeSpan godzRozpPracy = dane_Dni.godz_rozp_pracy;
            TimeSpan godzZakonczPracy = dane_Dni.godz_zakoncz_pracy;
            double godzNadlPlatne50 = (double)dane_Dni.Godz_nadl_platne_z_dod_50;
            double godzNadlPlatne100 = (double)dane_Dni.Godz_nadl_platne_z_dod_100;

            double czasPrzepracowany = 0;

            if (godzZakonczPracy < godzRozpPracy)
            {
                czasPrzepracowany = (TimeSpan.FromHours(24) - godzRozpPracy).TotalHours + godzZakonczPracy.TotalHours;
            }
            else
            {
                czasPrzepracowany = (godzZakonczPracy - godzRozpPracy).TotalHours;
            }

            double czasPodstawowy = czasPrzepracowany - (godzNadlPlatne50 + godzNadlPlatne100);

            TimeSpan startPodstawowy = godzRozpPracy;
            TimeSpan endPodstawowy = startPodstawowy + TimeSpan.FromHours(czasPodstawowy);

            TimeSpan startNadl50 = endPodstawowy;
            TimeSpan endNadl50 = startNadl50 + TimeSpan.FromHours(godzNadlPlatne50);

            TimeSpan startNadl100 = endNadl50;
            TimeSpan endNadl100 = startNadl100 + TimeSpan.FromHours(godzNadlPlatne100);

            return (new TimeSpan((int)startPodstawowy.TotalHours % 24, startPodstawowy.Minutes, startPodstawowy.Seconds),
                new TimeSpan((int)endPodstawowy.TotalHours % 24, endPodstawowy.Minutes, endPodstawowy.Seconds),
                new TimeSpan((int)startNadl50.TotalHours % 24, startNadl50.Minutes, startNadl50.Seconds),
                new TimeSpan((int)endNadl50.TotalHours % 24, endNadl50.Minutes, endNadl50.Seconds),
                new TimeSpan((int)startNadl100.TotalHours % 24, startNadl100.Minutes, startNadl100.Seconds),
                new TimeSpan((int)endNadl100.TotalHours % 24, endNadl100.Minutes, endNadl100.Seconds));
        }
    }
}
