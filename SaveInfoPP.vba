Sub Gerar()

 Dim objPPT As PowerPoint.Application
 Dim apresModelo As PowerPoint.Presentation
 Dim priSlide As PowerPoint.slide
 Dim formaSlide As PowerPoint.Shape
 Dim encontrouTXT As PowerPoint.TextRange
 Dim itemSubst As String, valsubst As String
 Dim j As Integer
 
 Const ppSaveAsPDF = 32

 Set objPPT = New PowerPoint.Application
 objPPT.Visible = True

 Set apresModelo = objPPT.Presentations.Open("M:\localização do ppt.pptx")
 Set priSlide = apresModelo.Slides(1)

 For Each formaSlide In priSlide.Shapes
 If formaSlide.HasTextFrame Then
 If formaSlide.TextFrame.HasText Then
 For j = 2 To 24
 itemSubst = Trim(Cells(1, j).Value)
 Dim ultimaLinhaTabela As Long
 ultimaLinhaTabela = Cells(Rows.Count, j).End(xlUp).Row
 valsubst = Cells(ultimaLinhaTabela, j).Value


 If Len(itemSubst) > 0 Then
 Set encontrouTXT = formaSlide.TextFrame.TextRange.Find(itemSubst)

 If Not (encontrouTXT Is Nothing) Then
 encontrouTXT.Text = valsubst
 Else
 Debug.Print "Texto não encontrado: " & itemSubst
 End If
 Else
 Debug.Print "ItemSubst vazio na coluna " & j
 End If
 Next j
 End If
 End If
 Next formaSlide

 Dim nomeArquivo As String
 Dim horarioPreenchimento As String

 horarioPreenchimento = Format(Now, "dd.mm.yyyy hh-mm")

 nomeArquivo = "M:\localização do arquivo - " & Cells(ultimaLinhaTabela, 2).Value & " - " & Cells(ultimaLinhaTabela, 2).Value & " (" & horarioPreenchimento & ").pdf"
 apresModelo.SaveAs nomeArquivo, ppSaveAsPDF

 objPPT.Quit

 Set objPPT = Nothing
 Set apresModelo = Nothing
 Set priSlide = Nothing
 Set encontrouTXT = Nothing

 MsgBox "Criado com sucesso!"

End Sub
