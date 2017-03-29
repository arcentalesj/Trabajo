USE [CSC]
GO

/****** Object:  Table [dbo].[cierre]    Script Date: 03/29/2017 09:22:57 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[cierre]') AND type in (N'U'))
DROP TABLE [dbo].[cierre]
GO

USE [CSC]
GO

/****** Object:  Table [dbo].[cierre]    Script Date: 03/29/2017 09:22:57 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[cierre](
	[cierre_equipo] [char](5) NULL,
	[cierre_fecha] [datetime] NULL,
	[cierre_mes] [int] NULL,
	[cierre_anio] [int] NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

