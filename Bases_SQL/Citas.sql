USE [CSC]
GO

/****** Object:  Table [dbo].[Citas]    Script Date: 03/29/2017 09:25:04 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Citas]') AND type in (N'U'))
DROP TABLE [dbo].[Citas]
GO

USE [CSC]
GO

/****** Object:  Table [dbo].[Citas]    Script Date: 03/29/2017 09:25:04 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Citas](
	[cita_cedula] [varchar](13) NOT NULL,
	[cita_fecha] [datetime] NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

