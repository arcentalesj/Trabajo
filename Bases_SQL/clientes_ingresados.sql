USE [CSC]
GO

/****** Object:  View [dbo].[clientes_ingresados]    Script Date: 03/29/2017 09:26:55 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[clientes_ingresados]'))
DROP VIEW [dbo].[clientes_ingresados]
GO

USE [CSC]
GO

/****** Object:  View [dbo].[clientes_ingresados]    Script Date: 03/29/2017 09:26:56 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[clientes_ingresados]
AS
SELECT DISTINCT 
                      TOP (100) PERCENT dbo.seguimiento.seg_cedula AS cedula, LTRIM(RTRIM(dbo.seguimiento.seg_apellido)) + ' ' + LTRIM(RTRIM(dbo.seguimiento.seg_nombre)) AS nombres, 
                      dbo.seguimiento.seg_feccaptacion AS Contactado, dbo.seguimiento.seg_feccontacto AS captado, dbo.asesores.nombre AS nomvende, dbo.seguimiento.seg_idvendedor AS idvendedor, 
                      dbo.articulos.dsmod AS Producto, dbo.seguimiento.seg_seguimiento AS segir, dbo.puntoventa.nom_puntoventa AS Ptoventa, dbo.seguimiento.seg_codarea1 AS codarea1, 
                      dbo.seguimiento.seg_Telefono AS telefono, dbo.seguimiento.seg_codarea2 AS codarea2, dbo.seguimiento.seg_telefono2 AS fono2, dbo.seguimiento.seg_buro1 AS buro, 
                      dbo.seguimiento.seg_fec_cita AS cita, dbo.seguimiento.Seg_celular AS celular, dbo.seguimiento.seg_fecevalua AS FechaEvaluado, dbo.seguimiento.seg_esta2 AS confirmaVenta, 
                      dbo.seguimiento.seg_fecesta2 AS FechaConfirmado, dbo.seguimiento.seg_actestado AS UsuarioVenta, dbo.observacc.obscc_detalle AS Observaciones
FROM         dbo.seguimiento INNER JOIN
                      dbo.asesores ON dbo.seguimiento.seg_idvendedor = dbo.asesores.codven INNER JOIN
                      dbo.articulos ON dbo.seguimiento.seg_articulo = dbo.articulos.cdart LEFT OUTER JOIN
                      dbo.puntoventa ON dbo.seguimiento.seg_IdPtoVnta = dbo.puntoventa.ID_puntoventa LEFT OUTER JOIN
                      dbo.observacc ON dbo.seguimiento.seg_observaevalua = dbo.observacc.obscc_codigo
WHERE     (dbo.seguimiento.seg_seguimiento IS NULL)
ORDER BY nombres

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[47] 4[22] 2[19] 3) )"
      End
      Begin PaneConfiguration = 1
         NumPanes = 3
         Configuration = "(H (1 [50] 4 [25] 3))"
      End
      Begin PaneConfiguration = 2
         NumPanes = 3
         Configuration = "(H (1[50] 2[25] 3) )"
      End
      Begin PaneConfiguration = 3
         NumPanes = 3
         Configuration = "(H (4 [30] 2 [40] 3))"
      End
      Begin PaneConfiguration = 4
         NumPanes = 2
         Configuration = "(H (1 [56] 3))"
      End
      Begin PaneConfiguration = 5
         NumPanes = 2
         Configuration = "(H (2 [66] 3))"
      End
      Begin PaneConfiguration = 6
         NumPanes = 2
         Configuration = "(H (4 [50] 3))"
      End
      Begin PaneConfiguration = 7
         NumPanes = 1
         Configuration = "(V (3))"
      End
      Begin PaneConfiguration = 8
         NumPanes = 3
         Configuration = "(H (1 [56] 4 [18] 2))"
      End
      Begin PaneConfiguration = 9
         NumPanes = 2
         Configuration = "(H (1 [75] 4))"
      End
      Begin PaneConfiguration = 10
         NumPanes = 2
         Configuration = "(H (1[66] 2) )"
      End
      Begin PaneConfiguration = 11
         NumPanes = 2
         Configuration = "(H (4 [60] 2))"
      End
      Begin PaneConfiguration = 12
         NumPanes = 1
         Configuration = "(H (1) )"
      End
      Begin PaneConfiguration = 13
         NumPanes = 1
         Configuration = "(V (4))"
      End
      Begin PaneConfiguration = 14
         NumPanes = 1
         Configuration = "(V (2))"
      End
      ActivePaneConfig = 0
   End
   Begin DiagramPane = 
      Begin Origin = 
         Top = 0
         Left = 0
      End
      Begin Tables = 
         Begin Table = "seguimiento"
            Begin Extent = 
               Top = 6
               Left = 38
               Bottom = 247
               Right = 212
            End
            DisplayFlags = 280
            TopColumn = 16
         End
         Begin Table = "articulos"
            Begin Extent = 
               Top = 97
               Left = 646
               Bottom = 205
               Right = 797
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "puntoventa"
            Begin Extent = 
               Top = 261
               Left = 472
               Bottom = 366
               Right = 670
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "observacc"
            Begin Extent = 
               Top = 151
               Left = 245
               Bottom = 259
               Right = 396
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "asesores"
            Begin Extent = 
               Top = 55
               Left = 409
               Bottom = 175
               Right = 607
            End
            DisplayFlags = 280
            TopColumn = 0
         End
      End
   End
   Begin SQLPane = 
   End
   Begin DataPane = 
      Begin ParameterDefaults = ""
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 3105
         Alias = 1470
         Table = 1170
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
  ' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'clientes_ingresados'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N'       Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'clientes_ingresados'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'clientes_ingresados'
GO

