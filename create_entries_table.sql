CREATE TABLE Entries (
    Entry_ID AUTOINCREMENT PRIMARY KEY,
    Employee_ID INTEGER,
    [Month] VARCHAR(30),
    [Year] VARCHAR(4),
    Contract_Name VARCHAR(40),
    Hours DOUBLE,
    Source_File VARCHAR(50),
    Sheet_Name VARCHAR(20)
);


