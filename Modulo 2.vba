Private Sub TextBox2_Change() 'Data Inicio'

TextBox2.MaxLength = 10

 Select Case TextBox2.SelStart
 Case Is = 2, 5
 TextBox2.SelText = "/"
 End Select

End Sub

Private Function CalcularDuracao() As String
 Dim dataHoraInicio As Date
 Dim dataHoraFim As Date
 Dim duracao As Double
 Dim dias As Long
 Dim horas As Long
 Dim minutos As Long

 ' Verifica se os campos têm valores válidos
 If IsDate(Me.TextBox2.Value & " " & Replace(Me.TextBox13.Value, "h", ":")) And _
 IsDate(Me.TextBox21.Value & " " & Replace(Me.TextBox14.Value, "h", ":")) Then

 ' Concatena data e hora e converte para tipo Date
 dataHoraInicio = CDate(Me.TextBox2.Value & " " & Replace(Me.TextBox13.Value, "h", ":"))
 dataHoraFim = CDate(Me.TextBox21.Value & " " & Replace(Me.TextBox14.Value, "h", ":"))

 ' Se a data final for menor, assume que passou para o dia seguinte
 If dataHoraFim < dataHoraInicio Then
 dataHoraFim = dataHoraFim + 1
 End If

 ' Calcula a diferença em dias (como número decimal)
 duracao = dataHoraFim - dataHoraInicio

 ' Extrai dias, horas e minutos
 dias = Int(duracao)
 horas = Int((duracao - dias) * 24)
 minutos = Round((((duracao - dias) * 24) - horas) * 60)

 ' Ajuste se minutos arredondarem para 60
 If minutos = 60 Then
 horas = horas + 1
 minutos = 0
 End If

 ' Ajuste se horas passarem de 24
 If horas = 24 Then
 dias = dias + 1
 horas = 0
 End If

 ' Monta o texto final
 CalcularDuracao = dias & " dia(s), " & horas & " horas e " & minutos & " minutos"
 Else
 CalcularDuracao = ""
 End If
End Function

Private Sub TextBox13_Change()
 Dim texto As String
 texto = Replace(TextBox13.Text, "h", "") ' Remove qualquer "h" digitado

 If Len(texto) = 4 Then
 TextBox13.Text = Left(texto, 2) & "h" & Right(texto, 2)
 TextBox13.SelStart = Len(TextBox13.Text)
 End If
End Sub

Private Sub ToggleButton1_Click()

 Dim ws As Worksheet
 Dim ultimaLinha As Long
 Dim ctrl As Control
 Dim camposVazios As Boolean
 camposVazios = False

 ' Verifica todos os TextBox do formulário
 For Each ctrl In Me.Controls
 If TypeName(ctrl) = "TextBox" Then
 If Trim(ctrl.Text) = "" Then
 camposVazios = True
 Exit For
 End If
 End If
 Next ctrl

 ' Verifica se os ListBox estão preenchidos
 If ListBox1.ListIndex = -1 Or ListBox2.ListIndex = -1 Or ListBox3.ListIndex = -1 Then
 camposVazios = True
 End If

 If camposVazios Then
 MsgBox "Por favor, preencha todos os campos antes de enviar.", vbExclamation, "Campos obrigatórios"
 Exit Sub
 End If

 Dim tbl As ListObject
 Set ws = ThisWorkbook.Sheets("Preenchimento")
 Set tbl = ws.ListObjects("Tabela4")

 Dim novaLinha As ListRow
 Set novaLinha = tbl.ListRows.Add

 Dim selecionados1 As String, selecionados2 As String, selecionados3 As String
 Dim selecionados4 As String, selecionados7 As String
 Dim contatosEnvolvidos As String, contatosImpactados As String
 Dim i As Integer, j As Integer

 ' Referência à Tabela2 na aba "Departamento"
 Dim wsDept As Worksheet
 Set wsDept = ThisWorkbook.Sheets("Departamento")

 With novaLinha
 ' Coletar selecionados do ListBox1
 For i = 0 To ListBox1.ListCount - 1
 If ListBox1.Selected(i) Then
 selecionados1 = selecionados1 & ListBox1.List(i) & ", "
 End If
 Next i

 ' Coletar selecionados do ListBox2 e buscar contatos
 For i = 0 To ListBox2.ListCount - 1
 If ListBox2.Selected(i) Then
 selecionados2 = selecionados2 & ListBox2.List(i) & ", "
 For j = 2 To 67 ' Tabela2 está em A2:B67
 If wsDept.Cells(j, 1).Value = ListBox2.List(i) Then
 If contatosEnvolvidos = "" Then
 contatosEnvolvidos = wsDept.Cells(j, 2).Value
 Else
 contatosEnvolvidos = contatosEnvolvidos & ", " & wsDept.Cells(j, 2).Value
 End If
 Exit For
 End If
 Next j
 End If
 Next i

 ' Coletar selecionados do ListBox3 e buscar contatos
 For i = 0 To ListBox3.ListCount - 1
 If ListBox3.Selected(i) Then
 selecionados3 = selecionados3 & ListBox3.List(i) & ", "
 For j = 2 To 67
 If wsDept.Cells(j, 1).Value = ListBox3.List(i) Then
 If contatosImpactados = "" Then
 contatosImpactados = wsDept.Cells(j, 2).Value
 Else
 contatosImpactados = contatosImpactados & ", " & wsDept.Cells(j, 2).Value
 End If
 Exit For
 End If
 Next j
 End If
 Next i

 ' Coletar selecionados do ListBox4
 For i = 0 To ListBox4.ListCount - 1
 If ListBox4.Selected(i) Then
 selecionados4 = selecionados4 & ListBox4.List(i) & ", "
 End If
 Next i

 ' Preencher os dados na nova linha da Tabela4
 .Range(1, 1).Value = TextBox1.Value ' Incidente
 .Range(1, 2).Value = TextBox2.Value ' Data
 .Range(1, 3).Value = CalcularDuracao ' Duração
 .Range(1, 4).Value = TextBox13.Value ' Início
 .Range(1, 5).Value = TextBox14.Value ' Fim
 .Range(1, 6).Value = selecionados1 ' Dimensão
 .Range(1, 7).Value = selecionados2 ' D. Env.
 .Range(1, 8).Value = contatosEnvolvidos ' Contatos D. Env.
 .Range(1, 9).Value = selecionados3 ' D. Imp.
 .Range(1, 10).Value = contatosImpactados ' Contatos D. Imp.
 .Range(1, 11).Value = TextBox11.Value ' Chave
 .Range(1, 12).Value = selecionados4 ' Conduzindo
 .Range(1, 13).Value = TextBox15.Value ' Nome
 .Range(1, 14).Value = TextBox4.Value ' Descr.
 .Range(1, 15).Value = TextBox16.Value ' Raiz
 .Range(1, 16).Value = TextBox5.Value ' Impactos
 .Range(1, 17).Value = TextBox17.Value ' Tomadas
 .Range(1, 18).Value = TextBox18.Value ' Plano
 .Range(1, 19).Value = TextBox6.Value ' Valor
 .Range(1, 20).Value = TextBox19.Value ' Acionado
 .Range(1, 21).Value = ListBox6.Value ' Aprovado
 .Range(1, 23).Value = TextBox20.Value ' Ponto focal
 End With

 ' Limpa os campos do formulário
 For Each ctrl In Me.Controls
 If TypeName(ctrl) = "TextBox" Then ctrl.Value = ""
 If TypeName(ctrl) = "ListBox" Then ctrl.ListIndex = -1
 Next ctrl

 MsgBox "Dados salvos com sucesso!", vbInformation, "Sucesso"

End Sub

    👏
    👍
    😊

Private Sub UserForm_Initialize()

 ' Configuração das ListBoxes
 ListBox1.RowSource = "Departamento!C2:C6"
 ListBox2.RowSource = "Departamento!A2:A67"
 ListBox3.RowSource = "Departamento!A2:A67"
 ListBox4.RowSource = "Pessoas!D2:D15"
 ListBox6.RowSource = "Pessoas!F2:F5"

End Sub
