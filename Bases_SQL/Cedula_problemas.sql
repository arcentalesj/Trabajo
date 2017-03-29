USE [CSC]
GO

/****** Object:  Table [dbo].[cedula_problemas]    Script Date: 03/29/2017 09:22:09 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[cedula_problemas](
	[ced_cedula] [varchar](13) NOT NULL,
	[ced_nombre] [varchar](30) NOT NULL,
	[ced_apellido] [varchar](30) NOT NULL,
	[ced_motivo] [varchar](200) NOT NULL,
	[ced_usrregistra] [varchar](50) NULL,
	[ced_fecregistro] [datetime] NOT NULL,
	[ced_usrhabilita] [varchar](50) NULL,
	[ced_fechabilita] [datetime] NULL,
	[ced_estatus] [char](1) NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

