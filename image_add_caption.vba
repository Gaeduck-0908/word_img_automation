Sub image_add_caption()
With Selection
.InsertCaption Label:="그림", TitleAutoText:="InsertCaption1", _
Title:="", Position:=wdCaptionPositionBelow, ExcludeLabel:=0
.Style = ActiveDocument.Styles("캡션[C1]")
.HomeKey Unit:=wdLine
.TypeText Text:="["
.EndKey Unit:=wdLine
.TypeText Text:="] 캡션작성"
.MoveUp Unit:=wdLine, Count:=1
.Style = ActiveDocument.Styles("본문[C1]")
End With
End Sub
