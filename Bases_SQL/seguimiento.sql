USE [CSC]
GO

/****** Object:  Table [dbo].[seguimiento]    Script Date: 03/29/2017 09:25:42 ******/
IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[seguimiento]') AND type in (N'U'))
DROP TABLE [dbo].[seguimiento]
GO

USE [CSC]
GO

/****** Object:  Table [dbo].[seguimiento]    Script Date: 03/29/2017 09:25:42 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[seguimiento](
	[seg_cedula] [varchar](13) NOT NULL,
	[seg_nombre] [varchar](30) NOT NULL,
	[seg_apellido] [varchar](30) NOT NULL,
	[seg_direccion] [varchar](80) NOT NULL,
	[seg_trabajo] [varchar](80) NOT NULL,
	[seg_mail] [varchar](40) NULL,
	[Seg_celular] [char](10) NOT NULL,
	[seg_codarea1] [char](2) NOT NULL,
	[seg_Telefono] [char](7) NOT NULL,
	[seg_codarea2] [char](2) NULL,
	[seg_telefono2] [char](7) NULL,
	[seg_civil] [char](3) NOT NULL,
	[seg_articulo] [char](2) NOT NULL,
	[seg_feccaptacion] [datetime] NOT NULL,
	[seg_feccontacto] [datetime] NULL,
	[seg_idvendedor] [int] NOT NULL,
	[seg_IdPtoVnta] [int] NOT NULL,
	[seg_tipodoc] [char](1) NULL,
	[seg_tomado] [char](1) NULL,
	[seg_buro1] [char](1) NULL,
	[seg_fecevalua] [datetime] NULL,
	[seg_fec_cita] [datetime] NULL,
	[seg_seguimiento] [char](2) NULL,
	[seg_anonac] [int] NULL,
	[seg_fecingdo] [datetime] NULL,
	[seg_actdatos] [varchar](50) NULL,
	[seg_fecdatos] [datetime] NULL,
	[seg_observaevalua] [char](3) NULL,
	[seg_feccontac] [datetime] NULL,
	[seg_fecestado] [datetime] NULL,
	[seg_actestado] [varchar](30) NULL,
	[seg_feceli] [datetime] NULL,
	[seg_elimina] [varchar](50) NULL,
	[seg_obselimina] [int] NULL,
	[seg_obselimope] [int] NULL,
	[seg_fecesta1] [datetime] NULL,
	[seg_esta2] [char](1) NULL,
	[seg_fecesta2] [datetime] NULL
) ON [PRIMARY]

GO

SET ANSI_PADDING OFF
GO

