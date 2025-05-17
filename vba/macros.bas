'COMPRIMRI IMAGENS
Sub ComprimirImagens()
    ' Applies formatting and opens the compress popup once for all pictures in the active sheet
    Dim s As Shape
    Dim shapeList As Collection
    Dim shapeNames() As String
    Dim i As Long

    Set shapeList = New Collection

    ' Collect all picture names from the active sheet
    For Each s In ActiveSheet.Shapes
        If s.Type = 13 Then ' msoPicture
            ' Apply image layout properties
            On Error Resume Next
            s.Placement = xlMoveAndSize
            s.PrintObject = True
            s.Locked = True
            s.LockAspectRatio = msoTrue
            On Error GoTo 0

            shapeList.Add s.Name
        End If
    Next s

    ' Select all pictures at once and show compress popup only once
    If shapeList.Count > 0 Then
        ReDim shapeNames(1 To shapeList.Count)
        For i = 1 To shapeList.Count
            shapeNames(i) = shapeList(i)
        Next i

        ActiveSheet.Shapes.Range(shapeNames).Select
        Application.CommandBars.ExecuteMso "PicturesCompress"
    Else
        MsgBox "No pictures found to compress on this sheet.", vbInformation
    End If
End Sub

' GERAR ENCOMENDAS: create suppliers
Sub GerarEncomendas()
    Dim resposta As VbMsgBoxResult
    Dim confirmar As VbMsgBoxResult
    Dim onlySupplier As String
    Dim pythonPath As String, scriptPath As String
    Dim command As String
    Dim objShell As Object
    Dim quote As String

    quote = Chr(34)

    ' Perguntar se é para gerar todas ou apenas uma
    resposta = MsgBox( _
        "Queres gerar TODAS as sheets de fornecedores?" & vbNewLine & _
        "Sim = todas; Não = só um", _
        vbQuestion + vbYesNoCancel + vbDefaultButton1, _
        "Gerar Encomendas" _
    )

    If resposta = vbCancel Then
        MsgBox "Operação cancelada.", vbInformation
        Exit Sub
    ElseIf resposta = vbYes Then
        onlySupplier = ""
    Else
        onlySupplier = InputBox("Indica o NOME do fornecedor a gerar:", "Fornecedor Único")
        If Trim(onlySupplier) = "" Then
            MsgBox "Operação cancelada.", vbExclamation
            Exit Sub
        End If
    End If

    ' Confirmação final antes de executar
    confirmar = MsgBox("Tens a certeza que queres continuar com a geração de fornecedores?", vbQuestion + vbYesNo, "Confirmar Ação")
    If confirmar = vbNo Then
        MsgBox "Operação cancelada.", vbInformation
        Exit Sub
    End If

    ' Caminhos
    pythonPath = "C:\Users\aimartins\AppData\Local\Programs\Python\Python312\python.exe"
    scriptPath = "C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\n116_repo\main.py"

    ' Construção do comando
    If onlySupplier = "" Then
        command = quote & pythonPath & quote & " " & _
                  quote & scriptPath & quote & " -m production create_supplier_sheets " & _
                  quote & "0" & quote
    Else
        command = quote & pythonPath & quote & " " & _
                  quote & scriptPath & quote & " -m production create_supplier_sheets " & _
                  quote & "0" & quote & " " & quote & onlySupplier & quote
    End If

    ' Executar
    Set objShell = VBA.CreateObject("WScript.Shell")
    objShell.Run command, 1, False
End Sub


'ADICIONAR ZONAS
Sub AdicionarZonas()
    Dim rowInput As String
    Dim zoneName As String
    Dim pythonPath As String, scriptPath As String
    Dim command As String
    Dim objShell As Object
    Dim quote As String

    quote = Chr(34)

    rowInput = InputBox("Indica a linha onde queres inserir a zona (ex: 15 ou 15,22):", "Linha da Zona")
    If Trim(rowInput) = "" Then
        MsgBox "Operação cancelada (sem linha).", vbExclamation
        Exit Sub
    End If

    zoneName = InputBox("Indica o nome da nova zona:", "Nome da Zona")
    If Trim(zoneName) = "" Then
        MsgBox "Operação cancelada (sem nome).", vbExclamation
        Exit Sub
    End If

    pythonPath = "C:\Users\aimartins\AppData\Local\Programs\Python\Python312\python.exe"
    scriptPath = "C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\n116_repo\main.py"

command = quote & pythonPath & quote & " " & _
          quote & scriptPath & quote & " -m production insert_zone_row " & _
          quote & rowInput & quote & " " & """dummy""" & " " & quote & zoneName & quote

    Set objShell = VBA.CreateObject("WScript.Shell")
    objShell.Run command, 1, False
End Sub

'ADICIONAR ARTIGOS
Sub AdicionarArtigos()
    Dim rowInput As String
    Dim pythonPath As String, scriptPath As String
    Dim command As String
    Dim objShell As Object
    Dim quote As String

    quote = Chr(34)

    rowInput = InputBox("Indica as linhas onde queres adicionar artigos (ex: 12,13):", "Inserir Artigos")

    If Trim(rowInput) = "" Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    pythonPath = "C:\Users\aimartins\AppData\Local\Programs\Python\Python312\pythonw.exe"
    scriptPath = "C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\n116_repo\main.py"

    command = quote & pythonPath & quote & " " & _
              quote & scriptPath & quote & " -m production insert_product_rows " & _
              quote & rowInput & quote

    Set objShell = VBA.CreateObject("WScript.Shell")
    objShell.Run command, 1, False
End Sub