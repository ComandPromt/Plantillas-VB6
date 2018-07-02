USE master
EXEC sp_addextendedproc 'xsp_UpdateSignalFile', 'XSP.dll'
GRANT EXECUTE ON xsp_UpdateSignalFile TO PUBLIC
GO

CREATE DATABASE Quotes
GO

USE Quotes
GO

CREATE TABLE Quotations
(
  Quotation varchar(256) NOT NULL, 
  Author    varchar(64)  NOT NULL
)
GO

INSERT INTO Quotations (Quotation, Author) VALUES ('Give me chastity and continence, but not yet.', 'Saint Augustine')
INSERT INTO Quotations (Quotation, Author) VALUES ('The use of COBOL cripples the mind; its teaching should therefore be regarded as a criminal offense.', 'Edsger Dijkstra')
INSERT INTO Quotations (Quotation, Author) VALUES ('C makes it easy to shoot yourself in the foot; C++ makes it harder, but when you do, it blows away your whole leg.', 'Bjarne Stroustrup')
INSERT INTO Quotations (Quotation, Author) VALUES ('A programmer is a device for turning coffee into code.', 'Jeff Prosise (with an assist from Paul Erdos)')
INSERT INTO Quotations (Quotation, Author) VALUES ('I have not failed. I''ve just found 10,000 ways that won''t work.', 'Thomas Edison')
INSERT INTO Quotations (Quotation, Author) VALUES ('Blessed is the man who, having nothing to say, abstains from giving wordy evidence of the fact.', 'George Eliot')
INSERT INTO Quotations (Quotation, Author) VALUES ('I think there is a world market for maybe five computers.', 'Thomas Watson')
INSERT INTO Quotations (Quotation, Author) VALUES ('Computers in the future may weigh no more than 1.5 tons.', 'Popular Mechanics')
INSERT INTO Quotations (Quotation, Author) VALUES ('I have traveled the length and breadth of this country and talked with the best people, and I can assure you that data processing is a fad that won''t last out the year.', 'Prentice-Hall business books editor')
INSERT INTO Quotations (Quotation, Author) VALUES ('640K ought to be enough for anybody.', 'William Gates III')
GO

CREATE TRIGGER DataChanged ON Quotations FOR INSERT, UPDATE, DELETE
AS EXEC master..xsp_UpdateSignalFile 'Quotes.Quotations'
GO