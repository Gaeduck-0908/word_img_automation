Sub Selected_image_w_board()
With Selection.InlineShapes(1) 
.LockAspectRatio = False
.Width = CentimetersToPoints(16.5)
With .Borders
.OutsideColor = wdColorBlack
.OutsideLineStyle = wdLineStyleSingle
.OutsideLineWidth = wdLineWidth075pt 
End With
End With
End Sub
