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
        private string Last_Mod_Osoba = "";
        private DateTime Last_Mod_Time = DateTime.Now;
        private string Optima_Connection_String = "";
        public void Set_Optima_ConnectionString(string NewConnectionString)
        {
            if (string.IsNullOrEmpty(NewConnectionString))
            {
                Program.error_logger.New_Custom_Error("Error: Empty Connection string in gv2");
                Console.WriteLine("Error: Empty Connection string in gv2");
                throw new Exception("Error: Empty Connection string in gv2");
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
                List <Nieobecnosci> ListNieobecnosci = Get_Nieobecności_Z_Grafiku(grafik);
                Dodaj_Dane_Do_Optimy(grafik, ListNieobecnosci);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        private List<Nieobecnosci> Get_Nieobecności_Z_Grafiku(Grafik grafik)
        {
            //get planowane nieobecnosci z grafiku
            List<Nieobecnosci> ListNieobecnosci = new();
            foreach (var dane_dni in grafik.dane_dni)
            {
                foreach (var dzien in dane_dni.dane_dnia)
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
            return ListNieobecnosci;
        }
        private void Get_Header_Karta_Info(IXLWorksheet worksheet, ref Grafik grafik)
        {
            var dane = worksheet.Cell(3, 1).GetValue<string>().Trim();
            dane = Regex.Replace(dane, @"\s{2,}", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Tytułu Grafiku", 3, 1, "Brak Tytułu Grafiku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            bool isParsed = int.TryParse(dane.Split(' ')[7], out int rok);
            if (!isParsed)
            {
                Program.error_logger.New_Error(dane, "Data Grafiku", 3, 1, "Błąd czytania daty");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            grafik.rok = rok;
            grafik.Set_Miesiac(dane.Split(' ')[6]);
            if(grafik.miesiac == 0)
            {
                Program.error_logger.New_Error(dane.Split(' ')[6], "Data Grafiku miesiac", 3, 1, "Błąd czytania miesiaca");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private void Get_Dane_Dni(IXLWorksheet worksheet, ref Grafik grafik)
        {
            CurrentPosition pozycja = new()
            {
                row = 6,
                col = 1
            };
            while (true)
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
                                    Program.error_logger.New_Error(kodzik, "Kod Aktywnosci Dnia", pozycja.col, pozycja.row, "Blędny kod aktywności dnia, wystąpił null w komórce");
                                    throw new Exception(Program.error_logger.Get_Error_String());
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
                catch
                {
                    throw;
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
                    if (poz.row > 100)
                    {
                        break;
                    }
                    var dane = worksheet.Cell(poz.row, poz.col).GetValue<string>().Trim();
                    if (!string.IsNullOrEmpty(dane))
                    {
                        idcounter++;
                        Legenda legenda = new();
                        legenda.kod = dane.Split('-')[0].Trim().Split(' ')[0].Trim();
                        if (legenda.kod == null)
                        {
                            legenda.kod = dane.Split('-')[0].Trim().Split('.')[0].Trim();
                        }
                        if (legenda.kod == null)
                        {
                            legenda.kod = dane.Split('-')[0].Trim();
                        }
                        if (legenda.kod.Contains('.'))
                        {
                            legenda.kod = legenda.kod.Split('.')[0].Trim();
                        }
                        if (legenda.kod == null)
                        {
                            Program.error_logger.New_Error(dane, "Linijka w legendie", poz.col, poz.row, "Program nie potrafi odczytać tej legendy. Wystąpił null. Zły format.");
                            var e = new Exception(Program.error_logger.Get_Error_String());
                            e.Data["kod"] = 69420;
                            throw e;
                        }
                        legenda.id = idcounter;
                        legenda.opis = dane;
                        grafik.legenda.Add(legenda);
                    }
                    poz.row++;
                }
                catch (Exception ex)
                {
                    if (ex.Data.Contains("kod") && ex.Data["kod"] is int kod && kod == 69420)
                    {
                        Program.error_logger.New_Custom_Error($"{ex.Message} W pliku {Program.error_logger.Nazwa_Pliku}, w zakładce {Program.error_logger.Nr_Zakladki}");
                        throw;
                    }
                    throw new Exception($"{ex.Message} W pliku {Program.error_logger.Nazwa_Pliku}, w zakładce {Program.error_logger.Nr_Zakladki}", ex);
                }
            }
        }
        private void Wpierdol_Plan_do_Optimy(Grafik grafik, SqlConnection connection, SqlTransaction tran)
        {
            foreach (var dane_Dni in grafik.dane_dni)
            {
                foreach (var dzien in dane_Dni.dane_dnia)
                {
                    try
                    {
                        using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertPlanDoOptimy, connection, tran))
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
                                        var e = new Exception($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        e.Data["Kod"] = 42069;
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
                                        var e = new Exception($"Error: Program nie rozpoznaje tego formatu czasu z legendy: {matchingLegenda.opis} w pliku {grafik.nazwa_pliku} z zakladki {grafik.nr_zakladki}. Nie wpisano dnia planu do bazy.");
                                        e.Data["Kod"] = 42069;
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
        }
        private void Wjeb_Nieobecnosci_do_Optimy(List<Nieobecnosci> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
        {
            List<List<Nieobecnosci>> Nieobecnosci = Podziel_Niobecnosci_Na_Osobne(ListaNieobecności);
            foreach (var ListaNieo in Nieobecnosci)
            {
                var dni_robocze = Ile_Dni_Roboczych(ListaNieo);
                var dni_calosc = ListaNieo.Count;
                try
                {
                    using (SqlCommand insertCmd = new SqlCommand(Program.sqlQueryInsertNieObecnoŚciDoOptimy, connection, tran))
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
                    connection.Close();
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