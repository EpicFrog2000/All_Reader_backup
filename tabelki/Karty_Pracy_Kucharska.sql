DROP TABLE Karta_Pracy_Kucharska_Pracownicy;
DROP TABLE Karty_Pracy_Kucharska;
DROP TABLE Karty_Pracy_Kucharska_Dane_Dni;
DROP TABLE Karty_Pracy_Kucharska_Legenda;
DROP TABLE Karty_Pracy_Kucharska_Usprawiedliwienia;

CREATE TABLE Karta_Pracy_Kucharska_Pracownicy(
    Id_Pracownika INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Imie VARCHAR(100) NOT NULL,
	Nazwisko VARCHAR(100) NOT NULL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karty_Pracy_Kucharska(
    Id_Karty INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Id_Pracownika INT,
	Oddzial VARCHAR(200),
	Zespol VARCHAR(200),
	Stanowisko VARCHAR(200),
	Nominal_Miesieczny_Ogolem INT,
	Miesiac TINYINT,
	Rok SMALLINT,
	Razem_Czas_Faktyczny_Przepracowany DECIMAL,
	Razem_Praca_WG_Grafiku DECIMAL,
	Razem_Przekr_Normy_Dobowej DECIMAL,
	Razem_Ilosc_Godzin_Z_Dodatkiem_50 DECIMAL,
	Razem_Ilosc_Godzin_Z_Dodatkiem_100 DECIMAL,
	Razem_Godziny_W_Niedziele DECIMAL,
	Razem_Godziny_W_Swieta DECIMAL,
	Razem_Godziny_W_Nocy DECIMAL,
	Razem_Dodatek_Szkodliwy_Ilosc_Godzin DECIMAL,
	Praca_Po_Absencji DECIMAL,
	Ogolem_Godziny_Nadliczbowe DECIMAL,
	Brak_Do_Nominalu DECIMAL,
	Spoznienia DECIMAL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karty_Pracy_Kucharska_Dane_Dni(
    Id_Karty INT,
	Dzien TINYINT,
	Godzina_Rozpoczêcia_Pracy TIME,
	Absencja VARCHAR(50),
	Godzina_Zakonczenia_Pracy TIME,
	Czas_Faktyczny_Przepracowany DECIMAL,
	Praca_WG_Grafiku DECIMAL,
	Przekr_Normy_Dobowej DECIMAL,
	Ilosc_Godzin_Z_Dodatkiem_50 DECIMAL,
	Ilosc_Godzin_Z_Dodatkiem_100 DECIMAL,
	Godziny_W_Niedziele DECIMAL,
	Godziny_W_Swieta DECIMAL,
	Godziny_W_Nocy DECIMAL,
	Dodatek_Szkodliwy_Ilosc_Godzin DECIMAL,
	Dodatek_Szkodliwy_Rodzaj_czynnosci VARCHAR(100),
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karty_Pracy_Kucharska_Legenda(
	Id_Kodu INT,
	Kod VARCHAR(10),
	Opis VARCHAR(200),
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Karty_Pracy_Kucharska_Usprawiedliwienia(
	ID_Karty INT,
	Opis VARCHAR(200),
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);