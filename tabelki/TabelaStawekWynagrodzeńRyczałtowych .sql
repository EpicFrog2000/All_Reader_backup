DROP TABLE Stawki_Wynagrodzen_Ryczaltowych_Detale;
DROP TABLE Stawki_Wynagrodzen_Ryczaltowych_Relacje;


CREATE TABLE Stawki_Wynagrodzen_Ryczaltowych_Relacje(
    Id_Relacji INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
    Nr_Relacji VARCHAR(10) NOT NULL,
    Opis_Relacji1 VARCHAR(200) NOT NULL,
	Opis_Relacji2 VARCHAR(200) NOT NULL,
	Rocznik VARCHAR(10) NOT NULL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL,
);

CREATE TABLE Stawki_Wynagrodzen_Ryczaltowych_Detale(
    Id_Relacji INT NOT NULL,
	Nr_Relacji_Detal VARCHAR(10),
	Opis VARCHAR(200),
	Czas_Relacji_Calkowity DECIMAL,
	Czas_Pracy_Ogolem DECIMAL,
	Czas_Pracy_Podstawowe DECIMAL,
	Godziny_Nadliczbowe_50 DECIMAL,
	Godziny_Nadliczbowe_100 DECIMAL,
	Godziny_Pracy_W_Nocy DECIMAL,
	Czas_Odpoczynku DECIMAL,
	Podstawowa_Stawka_Godzinowa DECIMAL,
	Podstawowe_Wynagrodzenie_Ryczaltowe DECIMAL,
	Wynagrodzenie_Za_Godziny_NadLiczbowe DECIMAL,
	Dodatek_Za_Prace_W_Nocy DECIMAL,
	Wynagrodzenie_Ryczaltowe_Calkowite DECIMAL,
	Dodatek_Wyjazdowy DECIMAL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);
