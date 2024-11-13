using All_Readeer;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml;

class Program
{
    private static string Files_Folder = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Wszystkie pliki";
    private static string Errors_File_Folder = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Errors\\";
    private static string Bad_Files_Folder = "G:\\ITEGER\\staż\\obecności\\All_Reader\\Bad Files\\";
    private static string Optima_Conection_String = "Server=ITEGER-NT;Database=CDN_Wars_Test_3_;User Id=sa;Password=cdn;Encrypt=True;TrustServerCertificate=True;";
    public static Error_Logger error_logger = new();
    public static void Main()
    {
        //Wpierdol do while(true){} jeśli to tyle
        ZrobToWieszCoNoWieszOCoMiChodzi();
    }

    public static void ZrobToWieszCoNoWieszOCoMiChodzi()
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
            File.WriteAllText(Errors_File_Folder+"Errors.txt", string.Empty);
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

        string[] filePaths = Directory.GetFiles(Files_Folder);

        if (filePaths.Length == 0) {
            Console.WriteLine("Nie znaleziono żadnych plików");
            return;
        }

        error_logger.Set_Error_File_Path(Errors_File_Folder);
        foreach (string filePath in filePaths)
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine($"Czytanie: {System.IO.Path.GetFileNameWithoutExtension(filePath)}");
            Console.ForegroundColor = ConsoleColor.White;
            try
            {
                var workbook = new XLWorkbook(filePath);
            }
            catch
            {
                Console.WriteLine($"Plik to nie arkusz xlsx: {filePath}.");
                Console.ReadLine();
                continue;
            }
            error_logger.Nazwa_Pliku = filePath;
            var (Last_Mod_Osoba, Last_Mod_Time) = Get_File_Meta_Info(filePath);
            if (Last_Mod_Osoba == "Error") { throw new Exception("Error reading file"); }
            int ilosc_zakladek = 0;
            using (var workbook = new XLWorkbook(filePath))
            {
                ilosc_zakladek = workbook.Worksheets.Count;
                for (int i = 1; i <= ilosc_zakladek; i++)
                {
                    error_logger.Nr_Zakladki = i;
                    var zakladka = workbook.Worksheet(i);
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
                            karta_Pracy_Reader_V2.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time);
                        }
                        catch
                        {
                            Cp_File_To_Bad_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                    else if (typ_pliku == 2)
                    {
                        try
                        {
                            Karta_Pracy_Reader karta_Pracy_Reader = new();
                            karta_Pracy_Reader.Set_Optima_ConnectionString(Optima_Conection_String);
                            karta_Pracy_Reader.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time);
                        }
                        catch
                        {
                            Cp_File_To_Bad_Files_Folder(filePath, error_logger.Nr_Zakladki);
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
                            Cp_File_To_Bad_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                    else if (typ_pliku == 3)
                    {
                        try
                        {
                            Grafik_Pracy_Reader_v2024 grafik_Pracy_Reader_v2024 = new();
                            grafik_Pracy_Reader_v2024.Set_Optima_ConnectionString(Optima_Conection_String);
                            grafik_Pracy_Reader_v2024.Process_Zakladka_For_Optima(zakladka, Last_Mod_Osoba, Last_Mod_Time);
                        }
                        catch
                        {
                            Cp_File_To_Bad_Files_Folder(filePath, error_logger.Nr_Zakladki);
                            continue;
                        }
                    }
                    // TODO dodać zwlonienia/urlopy z grafików i kartareaderv1
                    // TODO dodać support dla Zachód - zespół utrzymania czystości - Szczecin - karty pracy.xlsx bo obok siebie i pod są karty xdd
                    // grafik v2024 //TODO SPRAWDZ KILKA GRAFIKOW POD SOBĄ i sprawdz multiples of LICZBA GODZIN
                    // lepsze i wiecej errorów
                    // pewno bedzie wiecej jebanych kurwa edgecasów JAJEBE

                }
            }
            Console.ReadLine();
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
    private static void Cp_File_To_Bad_Files_Folder(string filePath, int sheetIndex)
    {
        try
        {
            string fileName = System.IO.Path.GetFileNameWithoutExtension(filePath) + "_copy.xlsx";
            string destPath = System.IO.Path.Combine(Bad_Files_Folder, fileName);

            using (var sourceWorkbook = new XLWorkbook(filePath))
            {
                if (sheetIndex >= 1 && sheetIndex <= sourceWorkbook.Worksheets.Count)
                {
                    var sourceSheet = sourceWorkbook.Worksheet(sheetIndex);

                    XLWorkbook destWorkbook;

                    if (File.Exists(destPath))
                    {
                        destWorkbook = new XLWorkbook(destPath);
                    }
                    else
                    {
                        destWorkbook = new XLWorkbook();
                    }

                    var destSheet = destWorkbook.Worksheets.Add(sourceSheet.Name);

                    foreach (var row in sourceSheet.Rows())
                    {
                        var destRow = destSheet.Row(row.RowNumber());
                        foreach (var cell in row.Cells())
                        {
                            try
                            {
                                // Check if cell and destination row are valid
                                if (cell != null && destRow != null)
                                {
                                    var destCell = destRow.Cell(cell.Address.ColumnNumber);
                                    if (cell.Address.Equals(destCell.Address))
                                    {
                                        continue; // Pomijamy, jeśli komórki są takie same
                                    }

                                    destCell.Value = cell.Value;

                                    // Copy styles only if they exist
                                    if (cell.Style != null)
                                    {
                                        // Font
                                        if (cell.Style.Font != null)
                                            destCell.Style.Font = cell.Style.Font;

                                        // Background Fill
                                        if (cell.Style.Fill != null)
                                            destCell.Style.Fill = cell.Style.Fill;

                                        // Alignment
                                        if (cell.Style.Alignment != null)
                                            destCell.Style.Alignment = cell.Style.Alignment;

                                        // Borders
                                        if (cell.Style.Border != null)
                                        {
                                            destCell.Style.Border.TopBorder = cell.Style.Border.TopBorder;
                                            destCell.Style.Border.BottomBorder = cell.Style.Border.BottomBorder;
                                            destCell.Style.Border.LeftBorder = cell.Style.Border.LeftBorder;
                                            destCell.Style.Border.RightBorder = cell.Style.Border.RightBorder;
                                            destCell.Style.Border.TopBorderColor = cell.Style.Border.TopBorderColor;
                                            destCell.Style.Border.BottomBorderColor = cell.Style.Border.BottomBorderColor;
                                            destCell.Style.Border.LeftBorderColor = cell.Style.Border.LeftBorderColor;
                                            destCell.Style.Border.RightBorderColor = cell.Style.Border.RightBorderColor;
                                        }
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing cell {cell.Address}: {ex.Message}");
                            }
                        }
                    }
                    destWorkbook.SaveAs(destPath);
                }
                else
                {
                    Console.WriteLine($"Arkusz o indeksie '{sheetIndex}' nie istnieje w pliku '{filePath}'.");
                }
            }
        }
        catch
        {
            Console.WriteLine($"Arkusz o indeksie '{sheetIndex}' w pliku '{filePath}' nie mógł być skopiowany, skopiowano cały plik.");
            try
            {
                string fileName = System.IO.Path.GetFileName(filePath);
                string destPath = System.IO.Path.Combine(Bad_Files_Folder, fileName);
                File.Copy(filePath, destPath, true);
                Console.WriteLine($"Plik skopiowany do folderu z błędami: {destPath}");
            }
            catch (Exception copyEx)
            {
                Console.WriteLine($"Błąd przy kopiowaniu pliku: {copyEx.Message}");
            }
        }
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