USE [CSC]
GO

/****** Object:  Table [dbo].[puntoventa]    Script Date: 03/29/2017 09:25:33 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[puntoventa]') AND type in (N'U'))
DROP TABLE [dbo].[puntoventa]
GO

USE [CSC]
GO

/****** Object:  Table [dbo].[puntoventa]    Script Date: 03/29/2017 09:25:33 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[puntoventa](
	[ID_puntoventa] [int] IDENTITY(1,1) NOT NULL,
	[nom_puntoventa] [varchar](50) NULL,
	[id_estatus] [char](1) NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

