
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, and Azure
-- --------------------------------------------------
-- Date Created: 12/08/2011 22:09:04
-- Generated from EDMX file: C:\Users\mdufour\Documents\MDWork\YahooAccess\YahooAccessData\StockYahooEntity.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [StockYahoo];
GO

-- Creating table 'StockSymbols'
CREATE TABLE [dbo].[StockSymbols] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [DateUpdate] datetime  NOT NULL,
    [Exchange] nvarchar(max)  NOT NULL,
    [Name] nvarchar(4000)  NOT NULL,
    [IndustryID] int  NOT NULL,
    [SectorID] int  NOT NULL,
    [Symbol] nvarchar(50)  NOT NULL,
    [SymbolNew] nvarchar(50)  NOT NULL,
    [StockID] int  NOT NULL
);
GO

-- Creating table 'StockErrors'
CREATE TABLE [dbo].[StockErrors] (
    [ID] int IDENTITY(1,1) NOT NULL,
    [DateUpdate] datetime  NOT NULL,
    [Symbol] nvarchar(50)  NOT NULL,
    [Description] nvarchar(max)  NOT NULL,
    [StockID] int  NOT NULL
);
GO

-- Creating primary key on [ID] in table 'StockSymbols'
ALTER TABLE [dbo].[StockSymbols]
ADD CONSTRAINT [PK_StockSymbols]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- Creating primary key on [ID] in table 'StockErrors'
ALTER TABLE [dbo].[StockErrors]
ADD CONSTRAINT [PK_StockErrors]
    PRIMARY KEY CLUSTERED ([ID] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- Creating foreign key on [StockID] in table 'StockErrors'
ALTER TABLE [dbo].[StockErrors]
ADD CONSTRAINT [FK_StockToStockError]
    FOREIGN KEY ([StockID])
    REFERENCES [dbo].[Stocks]
        ([ID])
    ON DELETE CASCADE ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_StockToStockError'
CREATE INDEX [IX_FK_StockToStockError]
ON [dbo].[StockErrors]
    ([StockID]);
GO

-- Creating foreign key on [StockID] in table 'StockSymbols'
ALTER TABLE [dbo].[StockSymbols]
ADD CONSTRAINT [FK_StockToStockSymbol]
    FOREIGN KEY ([StockID])
    REFERENCES [dbo].[Stocks]
        ([ID])
    ON DELETE CASCADE ON UPDATE NO ACTION;

-- Creating non-clustered index for FOREIGN KEY 'FK_StockToStockSymbol'
CREATE INDEX [IX_FK_StockToStockSymbol]
ON [dbo].[StockSymbols]
    ([StockID]);
GO

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------