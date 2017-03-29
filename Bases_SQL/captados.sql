USE [CSC]
GO

/****** Object:  View [dbo].[captados]    Script Date: 03/29/2017 09:26:30 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[captados]'))
DROP VIEW [dbo].[captados]
GO

USE [CSC]
GO

/****** Object:  View [dbo].[captados]    Script Date: 03/29/2017 09:26:30 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[captados]
AS
SELECT DISTINCT 
                      TOP (100) PERCENT dbo.seguimiento.seg_cedula AS cedula, dbo.seguimiento.seg_nombre AS nombre, dbo.seguimiento.seg_apellido AS apellido, dbo.seguimiento.seg_direccion AS direccion, 
                      dbo.seguimiento.Seg_celular AS celular, dbo.seguimiento.seg_mail AS correo, dbo.seguimiento.seg_codarea1 AS codarea1, dbo.seguimiento.seg_Telefono AS telefono, 
                      dbo.seguimiento.seg_codarea2 AS codarea2, dbo.seguimiento.seg_telefono2 AS fono2, dbo.seguimiento.seg_trabajo AS trabajo, dbo.seguimiento.seg_IdPtoVnta AS ptoventa, 
                      dbo.puntoventa.nom_puntoventa AS nomptoventa, dbo.seguimiento.seg_articulo AS idproducto, dbo.seguimiento.seg_articulo + ' - ' + Aplic.dbo.articulo.dsmod AS Producto, 
                      dbo.seguimiento.seg_buro1 AS buro, dbo.seguimiento.seg_feccontacto AS ingresado, dbo.seguimiento.seg_feccaptacion AS contactado, dbo.seguimiento.seg_seguimiento AS seguir, 
                      dbo.seguimiento.seg_fec_cita AS cita, dbo.seguimiento.seg_anonac AS nacer, dbo.seguimiento.seg_idvendedor AS idvendedor, dbo.asesores.nombre AS nombrevendedor, 
                      dbo.seguimiento.seg_civil + ' - ' + dbo.Estado.nom_estado AS Nomestado, YEAR(GETDATE()) - dbo.seguimiento.seg_anonac AS Edad, dbo.Estado.id_estado AS idestado
FROM         dbo.seguimiento INNER JOIN
                      Aplic.dbo.articulo ON dbo.seguimiento.seg_articulo = Aplic.dbo.articulo.cdart INNER JOIN
                      dbo.puntoventa ON dbo.seguimiento.seg_IdPtoVnta = dbo.puntoventa.ID_puntoventa INNER JOIN
                      dbo.asesores ON dbo.seguimiento.seg_idvendedor = dbo.asesores.codven INNER JOIN
                      dbo.Estado ON dbo.seguimiento.seg_civil = dbo.Estado.id_estado
WHERE     (dbo.seguimiento.seg_seguimiento IS NULL)
ORDER BY apellido, nombre

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[47] 4[32] 2[8] 3) )"
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
               Bottom = 310
               Right = 212
            End
            DisplayFlags = 280
            TopColumn = 7
         End
         Begin Table = "articulo (Aplic.dbo)"
            Begin Extent = 
               Top = 38
               Left = 673
               Bottom = 158
               Right = 871
            End
            DisplayFlags = 280
            TopColumn = 6
         End
         Begin Table = "puntoventa"
            Begin Extent = 
               Top = 6
               Left = 250
               Bottom = 111
               Right = 448
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "asesores"
            Begin Extent = 
               Top = 150
               Left = 723
               Bottom = 270
               Right = 921
            End
            DisplayFlags = 280
            TopColumn = 3
         End
         Begin Table = "Estado"
            Begin Extent = 
               Top = 7
               Left = 478
               Bottom = 97
               Right = 676
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
      Begin ColumnWidths = 9
         Width = 284
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
         Width = 1500
      End
   End
   Begin CriteriaPane = 
      Begin ColumnWidths = 11
         Column = 3375
         Alias' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'captados'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane2', @value=N' = 1530
         Table = 1890
         Output = 720
         Append = 1400
         NewValue = 1170
         SortType = 1350
         SortOrder = 1410
         GroupBy = 1350
         Filter = 1350
         Or = 1350
         Or = 1350
         Or = 1350
      End
   End
End
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'captados'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=2 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'captados'
GO

