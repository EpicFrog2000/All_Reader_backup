DROP TABLE Karta_Ewidecji_Konduktorzy;
DROP TABLE Karty_Ewidecji;
DROP TABLE Karta_Ewidencji;
DROP TABLE Karta_Ewidecji_Dni;
DROP TABLE Godziny_Pracy_Dnia;
DROP TABLE Godziny_Odpoczynku_Dnia;


CREATE TABLE Karta_Ewidecji_Konduktorzy(
    Id_Konduktora INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Nr_Sluzbowy INT,
	Stanowisko VARCHAR(100),
	Imie VARCHAR(100) NOT NULL,
	Nazwisko VARCHAR(100) NOT NULL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karty_Ewidecji(
	Id_Karty INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
    Id_Konduktora INT,
	Miesiac INT,
	Rok INT,
	Nominal_M_CA INT,
	Total_Godziny_Lacznie_Z_Odpoczynkiem DECIMAL,
	Total_Godziny_Pracy DECIMAL,
	Total_Liczba_Godzin_Relacji_Z_Odpoczynkiem DECIMAL,
	Total_Liczba_Godzin_Relacji_Pracy DECIMAL,
	Total_Liczba_Godzin_Nocnych DECIMAl,
	Total_Liczba_NadGodzin_Ogolem_50 DECIMAL,
	Total_Liczba_NadGodzin_Ogolem_100 DECIMAL,
	Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 DECIMAL,
	Total_Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 DECIMAL,
	Suma_Fodzin_Przepracowanych_Plus_Absencja DECIMAL,
	Nadgodziny_50 DECIMAL,
	Nadgodziny_100 DECIMAL,
	Total_Liczba_Godzin_Absencji DECIMAL,
	Total_Sprzedaz_Bilety_Zagranica DECIMAL,
	Total_Sprzedaz_Bilety_Kraj DECIMAL,
	Total_Sprzedaz_Bilety_Globalne DECIMAL,
	Total_Sprzedaz_Bilety_Wartosc_Towarow DECIMAL,
	Total_Liczba_Napojow_Awaryjnych DECIMAL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karta_Ewidencji(
	Id_Karty_Ewidencji INT,
	Numer_Relacji VARCHAR(50),
	Relacja VARCHAR(200),
	Nr_Pociagu VARCHAR(200),
	Liczba_Godzin_Relacji_Z_Odpoczynkiem DECIMAL,
	Liczba_Godzin_Relacji_Pracy DECIMAL,
	Liczba_Godzin_Nadliczbowych_W_Ryczalcie_50 DECIMAL,
	Liczba_Godzin_Nadliczbowych_W_Ryczalcie_100 DECIMAL,
	Wartosc_Biletow_Zagranica DECIMAL,
	Wartosc_Biletow_Kraj DECIMAL,
	Wartosc_Biletow_Globalne DECIMAL,
	Wartosc_Towarow DECIMAL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karta_Ewidecji_Dni(
	Id_Karta_Ewidecji_Dni INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Id_Karta_Ewidencji INT,
	Dzien INT,
	Godziny_Lacznie_Z_Odpoczynkiem DECIMAL,
	Godziny_Pracy DECIMAL,
	Liczba_Godzin_Nocnych DECIMAL,
	Liczba_NadGodzin_Ogolem_50 DECIMAL,
	Liczba_NadGodzin_Ogolem_100 DECIMAL,
	Nazwa_Absencji VARCHAR(50),
	Liczba_Godzin_Absencji DECIMAL,
	Liczba_Napojow_Awaryjnych DECIMAL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Godziny_Pracy_Dnia(
	Id_Karta_Ewidecji_Dni INT,
	Dzien INT,
	Godziny_Pracy_Od TIME,
	Godziny_Pracy_Do TIME,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Godziny_Odpoczynku_Dnia(
	Id_Karta_Ewidecji_Dni INT,
	Dzien INT,
	Godziny_Odpoczynku_Od TIME,
	Godziny_Odpoczynku_Do TIME,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);