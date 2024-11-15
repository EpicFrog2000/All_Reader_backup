namespace All_Readeer
{
    internal class Error_Logger
    {
        // Plik excel na którym obecnie wykonwywane są operacje
        public string Nazwa_Pliku = "";

        // Zakładka na której wystąpił błąd
        public int Nr_Zakladki = 0;

        // Obecna wartość pola z błędem
        private string Wartosc_Pola = "";

        // Nazwa tego co powinno znaleźć się w tym polu
        private string Poprawna_Wartosc_Pola = "";

        // Scierzka do pliku w którym maja być zapisywane błędy
        private string ErrorFilePath = "";

        // Kolumna w której wystąpił błąd
        public int Kolumna = -1;

        // Kolumna w której wystąpił błąd
        public int Rzad = -1;

        // Dodatkowa wiadomośc na koncu errora w pliku
        private string OptionalMsg = "";

        //Czas wykrycia błędu
        private DateTime Data_Czas_Wykrycia_Bledu;

        /// <summary>
        /// Wkłada wartości z parametrów do pól klasy i dodaje błąd do pliku z errorami.
        /// </summary>
        /// <param name="optionalmsg">Jeśli jest podany to dopisze na końcu wiadomosci błędu w pliku.</param>
        public void New_Error(string wartoscPola, string nazwaPola, int kolumna, int rzad, string? optionalmsg = null)
        {
            Poprawna_Wartosc_Pola = nazwaPola;
            Wartosc_Pola = wartoscPola;
            Kolumna = kolumna;
            Rzad = rzad;
            Data_Czas_Wykrycia_Bledu = DateTime.Now;
            if (!string.IsNullOrEmpty(optionalmsg))
            {
                OptionalMsg += $" Dodatkowa informacja: {optionalmsg}";
            }
            Append_Error_To_File();
        }
        /// <summary>
        /// Zwraca wiadomość jaką wpisało by do pliku z errorami.
        /// </summary>
        /// <returns>Zwraca wiadomość jaką wpisało by do pliku z errorami.</returns>
        public string Get_Error_String()
        {
            string Wiadomosc = @$"
-------------------------------------------------------------------------------
Wystąpił błąd w pliku: {Nazwa_Pliku}
Zakładka nr: {Nr_Zakladki}
Kolumna nr: {Kolumna}
Rząd nr: {Rzad}
Powinna znaleźć się wartość: {Poprawna_Wartosc_Pola}, a jest: {Wartosc_Pola}
Data_czas wykrycia: {Data_Czas_Wykrycia_Bledu}
";
            if (!string.IsNullOrEmpty(OptionalMsg))
            {
                Wiadomosc += OptionalMsg;
                OptionalMsg = string.Empty;
            }
            Wiadomosc += Environment.NewLine + "-------------------------------------------------------------------------------" + Environment.NewLine;
            return Wiadomosc;
        }
        /// <summary>
        /// Wpisuje do pliku z errorami wiadomość z parametru.
        /// </summary>
        public void New_Custom_Error(string Error_Msg)
        {
            Error_Msg = "-------------------------------------------------------------------------------" + Environment.NewLine + Error_Msg + Environment.NewLine + "-------------------------------------------------------------------------------" + Environment.NewLine;
            Append_Error_To_File(Error_Msg);
        }
        public void Set_Error_File_Path(string New_Error_File_Path)
        {
            ErrorFilePath = New_Error_File_Path;
        }
        private void Append_Error_To_File()
        {
            if (ErrorFilePath == "") { throw new Exception("ErrorLogger nie posiada właściwej scierzki do pliku Errors.txt"); }
            var ErrorsLogFile = Path.Combine(ErrorFilePath, "Errors.txt");
            if (!File.Exists(ErrorsLogFile))
            {
                File.Create(ErrorsLogFile).Dispose();
            }
            File.AppendAllText(ErrorsLogFile, Get_Error_String() + Environment.NewLine);
        }
        private void Append_Error_To_File(string Error_Msg)
        {
            if (ErrorFilePath == "") { throw new Exception("ErrorLogger nie posiada właściwej scierzki do pliku Errors.txt"); }
            var ErrorsLogFile = Path.Combine(ErrorFilePath, "Errors.txt");
            if (!File.Exists(ErrorsLogFile))
            {
                File.Create(ErrorsLogFile).Dispose();
            }
            File.AppendAllText(ErrorsLogFile, Error_Msg + Environment.NewLine);
        }
    }
}