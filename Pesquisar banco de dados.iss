Begin Dialog OpenDB 146,124,220,50,"Buscar banco de dados", .FuncaoOpenDB
  TextBox 6,11,159,14, .tbxNomeBase
  PushButton 172,11,33,14, "Abrir", .PushButton1
  Text 6,4,58,6, "Nome do banco", .Text1
End Dialog
Dim dlgTela As OpenDB
Dim vBase As String

Sub Main
Dim button As Integer

             	button = Dialog(dlgTela)
	Getfocus = dlgTela.tbxNomeBase

	Select Case button
		Case 1
			Call BuscaDB(dlgTela.tbxNomeBase)
	End Select 
End Sub

Function BuscaDB(pBase As String)
	vBase = pBase
	vBase = Replace(vBase, ".IMD", "") + ".IMD"
	
	If Dir(vBase) = vbNullString Then
		MsgBox "Base não encontrata!"
		Call Main()
	Else
		Client.OpenDataBase(vBase)
	End If
End Function

Function Replace(temp_string As String, temp_find As String, temp_replace As String) 
    Dim str_len, i As Integer
    Dim new_string As String
    Dim temp_char As String
    Dim temp_one_char As String
    Dim temp_no_chars As Integer
    
    new_string = "" 'set the new string
    str_len = Len(temp_string) 'length of the string to search in
    temp_no_chars = Len(temp_find) 'number of characters to search for
    
    For i = 1 To str_len
        temp_char = Mid(temp_string,i, temp_no_chars) 'get a partial string to compare
        Temp_one_char = Mid(temp_string,i, 1) 'get the character it is based on
        If temp_char = temp_find  Then 'validate if the partial string is what we are lookign for
            'if it is, replace it and skip i to take into account the no of characters
            temp_char = temp_replace 
            new_string = new_string & temp_replace 
            i = i + (temp_no_chars - 1)
        Else
            new_string = new_string & Temp_one_char 
        End If
        
    Next i
    
    replace = new_string
End Function
