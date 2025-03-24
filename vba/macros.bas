Sub AdicionarArtigo()
    Dim rowInput As String
    rowInput = InputBox("Indica as linhas onde queres adicionar artigos:", "Inserir Artigo")

    If rowInput = "" Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    Dim objShell As Object
    Set objShell = VBA.CreateObject("WScript.Shell")

    Dim pythonPath As String
    pythonPath = """C:\Users\aimartins\AppData\Local\Programs\Python\Python39\pythonw.exe"""

    Dim scriptPath As String
    scriptPath = """C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\n116_repo\main.py"""

    Dim command As String
    command = pythonPath & " " & scriptPath & " insert_product_between_columns """ & rowInput & """"

    objShell.Run command, 1, False
End Sub

'Botão ApagarArtigo
Sub ApagarArtigo()
    Dim rowInput As String
    rowInput = InputBox("Indica as linhas onde queres apagar artigos:", "Apagar Artigo")

    If rowInput = "" Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    Dim objShell As Object
    Set objShell = VBA.CreateObject("WScript.Shell")

    Dim pythonPath As String
    pythonPath = """C:\Users\aimartins\AppData\Local\Programs\Python\Python39\pythonw.exe"""

    Dim scriptPath As String
    scriptPath = """C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\n116_repo\main.py"""

    Dim command As String
    command = pythonPath & " " & scriptPath & " delete_between_columns """ & rowInput & """"

    objShell.Run command, 1, False
End Sub

Sub AdicionarZona()
    Dim rowInput As String
    rowInput = InputBox("Indica as linhas onde queres adicionar a zona:", "Adicionar Zona")

    If rowInput = "" Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    Dim zoneName As String
    zoneName = InputBox("Nome da zona (ex: Zona 1):", "Nome da Zona")

    If zoneName = "" Then
        MsgBox "Operação cancelada.", vbExclamation
        Exit Sub
    End If

    Dim objShell As Object
    Set objShell = VBA.CreateObject("WScript.Shell")

    Dim pythonPath As String
    pythonPath = """C:\Users\aimartins\AppData\Local\Programs\Python\Python39\pythonw.exe""" '

    Dim scriptPath As String
    scriptPath = """C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\n116_repo\main.py""" '

    ' Combine the full command
    Dim command As String
    command = pythonPath & " " & scriptPath & " add_zone """ & rowInput & """ """ & zoneName & """"

    ' Run the Python script
    objShell.Run command, 1, False
End Sub

