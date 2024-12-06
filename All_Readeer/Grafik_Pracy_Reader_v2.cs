using ClosedXML.Excel;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics.Metrics;
using System.Globalization;
using System.Text.RegularExpressions;

namespace All_Readeer
{
    internal static class Grafik_Pracy_Reader_v2
    {
        private class Pracownik
        {
            public string Imie { get; set; } = "";
            public string Nazwisko { get; set; } = "";
            public string Akronim { get; set; } = "-1";
        }
        private class Grafik
        {
            public string nazwa_pliku = "";
            public int nr_zakladki = 0;
            public int rok { get; set; } = 0;
            public int miesiac { get; set; } = 0;
            public List<Dane_Dni> dane_dni = [];
            public List<Legenda> legenda = [];
            public List<Nieobecnosci> ListaNieobecnosci= [];
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
        private class Current_Position
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
        public static void Process_Zakladka_For_Optima(IXLWorksheet worksheet)
        {
            try
            {
                var Pozycje = Find_Grafiki(worksheet);
                List<Grafik> Grafiki_W_Zakladce = new();
                foreach (var pozycja in Pozycje)
                {
                    Grafik grafik = new Grafik();
                    grafik.nazwa_pliku = Program.error_logger.Nazwa_Pliku;
                    grafik.nr_zakladki = Program.error_logger.Nr_Zakladki;
                    Get_Header_Karta_Info(pozycja , worksheet, ref grafik);
                    Get_Dane_Dni(pozycja, worksheet, ref grafik);
                    Get_Legenda(pozycja, worksheet, ref grafik);
                    grafik.ListaNieobecnosci = Get_Nieobecności_Z_Grafiku(grafik);
                    Grafiki_W_Zakladce.Add(grafik);
                }
                Dodaj_Dane_Do_Optimy(Grafiki_W_Zakladce);
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                throw;
            }
        }
        private static List<Nieobecnosci> Get_Nieobecności_Z_Grafiku(Grafik grafik)
        {
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
        private static void Get_Header_Karta_Info(Current_Position pozycja, IXLWorksheet worksheet, ref Grafik grafik)
        {
            var dane = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
            dane = Regex.Replace(dane, @"\s{2,}", " ");
            if (string.IsNullOrEmpty(dane))
            {
                Program.error_logger.New_Error(dane, "Tytułu Grafiku", pozycja.col, pozycja.row, "Brak Tytułu Grafiku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            if (int.TryParse(dane.Split(' ')[7], out int rok))
            {
                grafik.rok = rok;
                grafik.Set_Miesiac(dane.Split(' ')[6]);
            }
            else
            {
                Program.error_logger.New_Error(dane, "Data Grafiku", pozycja.col, pozycja.row, "Niepoprawny format daty. Powinna być data w formacie np. '30.12.2024'");
                throw new Exception(Program.error_logger.Get_Error_String());
            }

            if(grafik.miesiac == 0)
            {
                Program.error_logger.New_Error(dane.Split(' ')[6], "Data Grafiku miesiac", pozycja.col, pozycja.row, "Błąd czytania miesiaca");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
            if (grafik.rok == 0)
            {
                Program.error_logger.New_Error(dane.Split(' ')[6], "Data Grafiku miesiac", pozycja.col, pozycja.row, "Błąd czytania roku");
                throw new Exception(Program.error_logger.Get_Error_String());
            }
        }
        private static List<Current_Position> Find_Grafiki(IXLWorksheet worksheet)
        {
            List<Current_Position> Pozycje = new();
            int Limiter = 1000;
            int counter = 0;
            foreach (var cell in worksheet.CellsUsed())
            {


                try
                {
                    if (cell.HasFormula && !cell.Address.ToString()!.Equals(cell.FormulaA1))
                    {
                        counter++;
                        if (counter > Limiter)
                        {
                            break;
                        }
                        continue;
                    }

                    if (cell.Value.ToString().Contains("GRAFIK PRACY"))
                    {
                        Pozycje.Add(new Current_Position()
                        {
                            row = cell.Address.RowNumber,
                            col = cell.Address.ColumnNumber
                        });
                    }
                }
                catch
                {
                    continue;
                }
            }
            return Pozycje;
        }
        private static void Get_Dane_Dni(Current_Position pozycja, IXLWorksheet worksheet, ref Grafik grafik)
        {
            pozycja.row += 3;
            while (true)
            {

                pozycja.col = 1;
                var nazwiskoiimie = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim().Replace("  ", " ");

                // jeśli 3 next row puste to wypierdalaj
                if(string.IsNullOrEmpty(nazwiskoiimie) && string.IsNullOrEmpty(worksheet.Cell(pozycja.row + 1, pozycja.col).GetFormattedString().Trim()) && string.IsNullOrEmpty(worksheet.Cell(pozycja.row + 2, pozycja.col).GetFormattedString().Trim())){
                    break;
                }
                if (!string.IsNullOrEmpty(nazwiskoiimie.Trim()) && nazwiskoiimie.Trim().Split(' ').Length < 3)
                {
                    Dane_Dni dane_dni = new();
                    try
                    {
                        dane_dni.pracownik.Nazwisko = nazwiskoiimie.Split(' ')[0].Trim();
                        dane_dni.pracownik.Imie = nazwiskoiimie.Split(' ')[1].Trim();
                        dane_dni.pracownik.Nazwisko = dane_dni.pracownik.Nazwisko.ToLower();
                        dane_dni.pracownik.Nazwisko = char.ToUpper(dane_dni.pracownik.Nazwisko[0], CultureInfo.CurrentCulture) + dane_dni.pracownik.Nazwisko.Substring(1);
                        dane_dni.pracownik.Imie = dane_dni.pracownik.Imie.ToLower();
                        dane_dni.pracownik.Imie = char.ToUpper(dane_dni.pracownik.Imie[0], CultureInfo.CurrentCulture) + dane_dni.pracownik.Imie.Substring(1);
                    }
                    catch
                    {
                        //Program.error_logger.New_Error(nazwiskoiimie, "Nazwisko Imie", pozycja.col, pozycja.row, "Niepoprawnie wpisane nazwisko i ime. Powinno być w formacie np. 'Nazwisko Imie'");
                        //throw new Exception(Program.error_logger.Get_Error_String());
                    }


                    pozycja.col = 3;
                    var nrDniaLubAkronim = worksheet.Cell(5, pozycja.col).GetFormattedString().Trim();
                    if (!string.IsNullOrEmpty(nrDniaLubAkronim) && nrDniaLubAkronim.ToLower().Contains("akronim"))
                    {
                        var akr = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim().Replace("  ", " ");
                        if (!string.IsNullOrEmpty(akr))
                        {
                            dane_dni.pracownik.Akronim = akr;
                        }
                        pozycja.col++;
                    }

                    while (true)
                    {
                        Dane_Dnia dane_dnia = new();
                        var nrDnia = worksheet.Cell(5, pozycja.col).GetFormattedString().Trim();
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
                        var kodzik = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
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
        }
        private static void Get_Legenda(Current_Position pozycja, IXLWorksheet worksheet, ref Grafik grafik)
        {
            int idcounter = 0;
            pozycja.row += 18;
            pozycja.col += 3;
            while (true)
            {
                try
                {
                    if (pozycja.row > 100)
                    {
                        break;
                    }
                    var dane = worksheet.Cell(pozycja.row, pozycja.col).GetFormattedString().Trim();
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
                            Program.error_logger.New_Error(dane, "Linijka w legendie", pozycja.col, pozycja.row, "Program nie potrafi odczytać tej legendy. Wystąpił null. Zły format.");
                            var e = new Exception(Program.error_logger.Get_Error_String());
                            e.Data["kod"] = 69420;
                            throw e;
                        }
                        legenda.id = idcounter;
                        legenda.opis = dane;
                        grafik.legenda.Add(legenda);
                    }
                    pozycja.row++;
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
        private static void Dodaj_Plan_do_Optimy(Grafik grafik, SqlConnection connection, SqlTransaction tran)
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
                            insertCmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(dane_Dni.pracownik));
                            if (Program.error_logger.Last_Mod_Osoba.Length > 20)
                            {
                                insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 20));
                            }
                            else
                            {
                                insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba);
                            }
                            if (Program.error_logger.Last_Mod_Osoba.Length > 50)
                            {
                                insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 50));
                            }
                            else
                            {
                                insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba);
                            }
                            insertCmd.Parameters.AddWithValue("@DataMod", Program.error_logger.Last_Mod_Time);
                            insertCmd.ExecuteScalar();
                        }
                    }
                    catch (SqlException ex)
                    {
                        tran.Rollback();
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
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
                        Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                        throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    }
                }
            }
        }
        private static void Dodaj_Nieobecnosci_do_Optimy(List<Nieobecnosci> ListaNieobecności, SqlTransaction tran, SqlConnection connection)
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
                        insertCmd.Parameters.AddWithValue("@PRI_PraId", Get_ID_Pracownika(ListaNieo[0].pracownik));
                        insertCmd.Parameters.AddWithValue("@NazwaNieobecnosci", "Nieobecność (B2B) (plan)");
                        insertCmd.Parameters.AddWithValue("@DniPracy", dni_robocze);
                        insertCmd.Parameters.AddWithValue("@DniKalendarzowe", dni_calosc);
                        insertCmd.Parameters.AddWithValue("@Przyczyna", 9); // "Nieobecności inne"
                        insertCmd.Parameters.Add("@DataOd", SqlDbType.DateTime).Value = dataniobecnoscistart;
                        insertCmd.Parameters.Add("@BaseDate", SqlDbType.DateTime).Value = dataBazowa;
                        insertCmd.Parameters.Add("@DataDo", SqlDbType.DateTime).Value = dataniobecnosciend;
                        if (Program.error_logger.Last_Mod_Osoba.Length > 20)
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 20));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@ImieMod", Program.error_logger.Last_Mod_Osoba);
                        }
                        if (Program.error_logger.Last_Mod_Osoba.Length > 50)
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba.Substring(0, 50));
                        }
                        else
                        {
                            insertCmd.Parameters.AddWithValue("@NazwiskoMod", Program.error_logger.Last_Mod_Osoba);
                        }
                        insertCmd.Parameters.AddWithValue("@DataMod", Program.error_logger.Last_Mod_Time);
                        insertCmd.ExecuteScalar();
                    }
                }
                catch (SqlException ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
                catch (FormatException)
                {
                    continue;
                }
                catch (Exception ex)
                {
                    tran.Rollback();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
            }

        }
        private static void Dodaj_Dane_Do_Optimy(List<Grafik>ListaGrafików)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
                {
                    connection.Open();
                    SqlTransaction tran = connection.BeginTransaction();
                    foreach (var grafik in ListaGrafików)
                    {
                        Dodaj_Nieobecnosci_do_Optimy(grafik.ListaNieobecnosci, tran, connection);
                        Dodaj_Plan_do_Optimy(grafik, connection, tran);
                    }
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Poprawnie dodawno planowane nieobecnosci z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.WriteLine($"Poprawnie dodawno plan z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki);
                    Console.ForegroundColor = ConsoleColor.White;
                    tran.Commit();
                    connection.Close();
                }
            }
            catch
            {
                throw;
            }
        }
        private static int Ile_Dni_Roboczych(List<Nieobecnosci> listaNieobecnosci)
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
        private static List<List<Nieobecnosci>> Podziel_Niobecnosci_Na_Osobne(List<Nieobecnosci> listaNieobecnosci)
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
        private static int Get_ID_Pracownika(Pracownik pracownik)
        {
            using (SqlConnection connection = new SqlConnection(Program.Optima_Conection_String))
            {
                try
                {
                    connection.Open();
                    using (SqlCommand getCmd = new SqlCommand(Program.sqlQueryGetPRI_PraId, connection))
                    {
                        getCmd.Parameters.AddWithValue("@Akronim ", pracownik.Akronim);
                        getCmd.Parameters.AddWithValue("@PracownikImieInsert", pracownik.Imie);
                        getCmd.Parameters.AddWithValue("@PracownikNazwiskoInsert", pracownik.Nazwisko);
                        var result = getCmd.ExecuteScalar();
                        if (result != null)
                        {
                            return Convert.ToInt32(result);
                        }
                    }
                    connection.Close();
                }
                catch (Exception ex)
                {
                    connection.Close();
                    Program.error_logger.New_Custom_Error(ex.Message + " z pliku: " + Program.error_logger.Nazwa_Pliku + " z zakladki: " + Program.error_logger.Nr_Zakladki + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                    throw new Exception(ex.Message + $" w pliku {Program.error_logger.Nazwa_Pliku} z zakladki {Program.error_logger.Nr_Zakladki}" + " nazwa zakladki: " + Program.error_logger.Nazwa_Zakladki);
                }
            }
            return 0;
        }
    }
}