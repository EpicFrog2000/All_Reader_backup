DROP TABLE Grafik_Pracy_Detale_Dni;
DROP TABLE Grafik_Pracy_Detale;
DROP TABLE Grafik_Pracy_Legenda;
DROP TABLE Grafik_Pracy_Pracownicy;
DROP TABLE Grafik_Pracy_Oddzialy;
DROP TABLE Grafiki_Pracy;
DROP TABLE Grafik_Pracy_Detale_Dni;
DROP TABLE Grafik_Pracy_Detale;
DROP TABLE Grafik_Pracy_Legenda;
DROP TABLE Grafik_Pracy_Pracownicy;
DROP TABLE Grafik_Pracy_Oddzialy;
DROP TABLE Grafiki_Pracy;



CREATE TABLE Grafik_Pracy_Oddzialy(
    Id_Oddzialu INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Nazwa VARCHAR(200) NOT NULL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Grafik_Pracy_Pracownicy(
    Id_Pracownika INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Imie VARCHAR(100) NOT NULL,
	Nazwisko VARCHAR(100) NOT NULL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL
);

CREATE TABLE Grafiki_Pracy(
    Id_Grafiku INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	Miesiac TINYINT,
	Rok SMALLINT,
	Nominal_Godzin TINYINT,
	Id_Oddzialu INT,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL,
	FOREIGN KEY (Id_Oddzialu) REFERENCES Grafik_Pracy_Oddzialy(Id_Oddzialu)
);

CREATE TABLE Grafik_Pracy_Legenda(
	Id_Grafiku INT,
	Id_Kodu INT,
	Opis VARCHAR(255),
	Ilosc_Godzin DECIMAL,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL,
	FOREIGN KEY (Id_Grafiku) REFERENCES Grafiki_Pracy(Id_Grafiku),
);

CREATE TABLE Grafik_Pracy_Detale(
    Id_Detalu INT NOT NULL  IDENTITY(1,1) PRIMARY KEY,
	Id_Grafiku INT NOT NULL,
	Id_Pracownika INT,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL,
	FOREIGN KEY (Id_Pracownika) REFERENCES Grafik_Pracy_Pracownicy(Id_Pracownika),
	FOREIGN KEY (Id_Grafiku) REFERENCES Grafiki_Pracy(Id_Grafiku),
);

CREATE TABLE Grafik_Pracy_Detale_Dni(
	Id_Detalu INT,
	Dzien TINYINT,
	Id_Kodu INT,
	Ostatnia_Modyfikacja_Data DATETIME NOT NULL,
	Ostatnia_Modyfikacja_Osoba VARCHAR(100) NOT NULL,
	FOREIGN KEY (Id_Detalu) REFERENCES Grafik_Pracy_Detale(Id_Detalu),
);


