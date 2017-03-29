USE [CSC]
GO

/****** Object:  View [dbo].[citas_del_dia]    Script Date: 03/29/2017 09:26:45 ******/
IF  EXISTS (SELECT * FROM sys.views WHERE object_id = OBJECT_ID(N'[dbo].[citas_del_dia]'))
DROP VIEW [dbo].[citas_del_dia]
GO

USE [CSC]
GO

/****** Object:  View [dbo].[citas_del_dia]    Script Date: 03/29/2017 09:26:45 ******/
SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

CREATE VIEW [dbo].[citas_del_dia]
AS
SELECT     TOP (100) PERCENT dbo.seguimiento.seg_cedula AS cedula, dbo.seguimiento.seg_nombre AS nombre, dbo.seguimiento.seg_apellido AS apellido, dbo.seguimiento.seg_trabajo AS trabajo, 
                      dbo.seguimiento.seg_direccion AS direccion, dbo.seguimiento.Seg_celular AS celular, dbo.seguimiento.seg_codarea1 AS codarea1, dbo.seguimiento.seg_Telefono AS telefono, 
                      dbo.seguimiento.seg_codarea2 AS codarea2, dbo.seguimiento.seg_telefono2 AS fono2, dbo.seguimiento.seg_mail AS correo, dbo.seguimiento.seg_idvendedor AS idvendedor, 
                      dbo.asesores.nombre AS nomvende, dbo.seguimiento.seg_buro1 AS buro, dbo.seguimiento.seg_articulo AS idProducto, Aplic.dbo.articulo.dsmod AS Producto, Aplic.dbo.articulo.costo AS precio, 
                      dbo.puntoventa.nom_puntoventa AS ptoventa, dbo.seguimiento.seg_feccontacto AS captado, dbo.seguimiento.seg_feccaptacion AS Contactado, dbo.seguimiento.seg_fec_cita AS Cita, 
                      dbo.seguimiento.seg_seguimiento, dbo.asesores.dsreg AS desage, dbo.seguimiento.seg_apellido + '  ' + dbo.seguimiento.seg_nombre AS Nombres
FROM         dbo.seguimiento INNER JOIN
                      dbo.puntoventa ON dbo.seguimiento.seg_IdPtoVnta = dbo.puntoventa.ID_puntoventa FULL OUTER JOIN
                      Aplic.dbo.articulo ON dbo.seguimiento.seg_articulo = Aplic.dbo.articulo.cdart LEFT OUTER JOIN
                      dbo.asesores ON dbo.seguimiento.seg_idvendedor = dbo.asesores.codven
WHERE     (LEN(dbo.seguimiento.seg_idvendedor) > 0) AND (NOT (dbo.seguimiento.seg_fec_cita IS NULL)) AND (dbo.seguimiento.seg_seguimiento IS NULL) AND (LEN(dbo.asesores.dsreg) > 0)
ORDER BY apellido, nombre

GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPane1', @value=N'[0E232FF0-B466-11cf-A24F-00AA00A3EFFF, 1.00]
Begin DesignProperties = 
   Begin PaneConfigurations = 
      Begin PaneConfiguration = 0
         NumPanes = 4
         Configuration = "(H (1[40] 4[20] 2[20] 3) )"
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
               Bottom = 282
               Right = 212
            End
            DisplayFlags = 280
            TopColumn = 8
         End
         Begin Table = "puntoventa"
            Begin Extent = 
               Top = 81
               Left = 270
               Bottom = 186
               Right = 468
            End
            DisplayFlags = 280
            TopColumn = 0
         End
         Begin Table = "articulo (Aplic.dbo)"
            Begin Extent = 
               Top = 6
               Left = 686
               Bottom = 126
               Right = 884
            End
            DisplayFlags = 280
            TopColumn = 7
         End
         Begin Table = "asesores"
            Begin Extent = 
               Top = 83
               Left = 482
               Bottom = 275
               Right = 680
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
         Column = 2625
         Alias = 2505
         Table = 2250
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
' , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'citas_del_dia'
GO

EXEC sys.sp_addextendedproperty @name=N'MS_DiagramPaneCount', @value=1 , @level0type=N'SCHEMA',@level0name=N'dbo', @level1type=N'VIEW',@level1name=N'citas_del_dia'
GO

