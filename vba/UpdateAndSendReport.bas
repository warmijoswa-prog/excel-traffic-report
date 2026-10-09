Attribute VB_Name = "Módulo1"
Sub Actualizar()
Attribute Actualizar.VB_ProcData.VB_Invoke_Func = " \n14"
'FORMULAS
    Sheets("ACTUALIZAR").Select
    Range("B1").Select
    ActiveCell.FormulaR1C1 = "=TEXT(NOW()-23,""DD/MM/YYYY"")"
    Range("B3").Select
    ActiveCell.FormulaR1C1 = _
        "=TEXT((NOW()-(-1+WEEKDAY(NOW(),2))),""dd"")&""/""&TEXT((NOW()-(-1+WEEKDAY(NOW(),2))),""mm"")&""/""&TEXT((NOW()-(-1+WEEKDAY(NOW(),2))),""yyyy"")"
    Range("A4").Select
    ActiveCell.FormulaR1C1 = _
        "=TEXT(NOW(),""dd"")&"" de ""&TEXT(NOW(),""mmmm"")&"" ""&TEXT(NOW(),""yyyy"")"
    Range("A5").Select
    ActiveCell.FormulaR1C1 = _
        "="" ""&TEXT(NOW()-1,""yyyy"")&TEXT(NOW()-1,""mm"")&TEXT(NOW()-1,""dd"")"
'ACTUALIZAR DATA
    ActiveWorkbook.RefreshAll
    Application.CalculateUntilAsyncQueriesDone
    
'ACTUALIZAR Y GUARDAR PPT
Dim fecha As String
Dim nombre_archivo As String
Dim ppt As Object
Set ppt = CreateObject("PowerPoint.Application")

nombre_archivo = Dir("\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\" & "*.*")
fecha = Sheets("ACTUALIZAR").Range("B5").Value
Name "\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\" & nombre_archivo As "\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\Informe trafico red Ecuador " & fecha & ".pptx"
nombre_archivo = Dir("\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\" & "*.*")

ppt.Visible = True
ppt.Presentations.Open "\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\" & nombre_archivo
ppt.ActivePresentation.UpdateLinks
ppt.ActivePresentation.SaveAs "\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\" & nombre_archivo
ppt.Quit
Set ppt = Nothing


'CORREO
    Dim cuerpo As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Set cuerpo = Nothing
    
    
    Sheets("CORREO").Select
    Range("A5:K23").Select
    Set cuerpo = Selection.SpecialCells(xlCellTypeVisible)
    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)
'++++++++++++++++INICIO MAIL++++++++++++++++++++++++++++++++
With OutMail
.Display
.To = Sheets("CORREO").Range("B1").Value
.CC = Sheets("CORREO").Range("B2").Value
.Subject = Sheets("CORREO").Range("B3").Value
.htmlBody = RangetoHTML(cuerpo) '& .htmlBody
'.Importance = 2
.Attachments.Add "\\fileuio03\Ingenieria\Calidad Servicio de Red\INFORME TRAFICO RED ECUADOR\PPT\" & nombre_archivo
End With
'+++++++++++++++++FIN MAIL+++++++++++++++++++++++++++++++
Range("A1").Select

Sheets("ACTUALIZAR").Select
Range("A1").Select

End Sub




Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

