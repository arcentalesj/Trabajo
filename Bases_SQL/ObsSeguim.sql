USE [CSC]
GO

/****** Object:  Table [dbo].[ObsSeguim]    Script Date: 03/29/2017 09:25:22 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[ObsSeguim]') AND type in (N'U'))
DROP TABLE [dbo].[ObsSeguim]
GO

USE [CSC]
GO

/****** Object:  Table [dbo].[ObsSeguim]    Script Date: 03/29/2017 09:25:22 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[ObsSeguim](
	[obs_cedula] [varchar](13) NULL,
	[obs_fecha] [datetime] NULL,
	[obs_Observa] [varchar](200) NULL,
	[obs_ingresa] [varchar](50) NULL,
	[obs_evalua] [int] NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

