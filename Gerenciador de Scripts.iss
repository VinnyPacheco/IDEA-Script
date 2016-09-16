Dim ListBox2$() AS string

Begin Dialog GerarScriptsDlg 52,24,330,178,"Gerenciador de Scripts - Fernandes & Fernandes", .funMainMenu
  ListBox 12,43,121,93, .ListBox1
  ListBox 170,43,121,93, ListBox2$(), .ListBox2
  PushButton 141,53,20,14, ">", .AdicionarBtn
  PushButton 141,97,20,14, "<", .RetirarBtn
  PushButton 295,70,20,14, "^", .SubirBtn
  PushButton 295,90,20,14, "v", .DescerBtn
  Text 12,35,50,8, "Scripts guardados", .LblScriptGuardado
  Text 170,35,40,8, "Scripts a gerar", .LblScriptGerar
  PushButton 251,143,40,12, "Gerar", .GerarBtn
  PushButton 12,143,40,12, "Sair", .SairBtn
  PushButton 126,143,50,12, "Alterar Diretório", .DiretorioBtn
  PushButton 141,70,20,14, ">>", .AdicionarTodosBtn
  PushButton 141,114,20,14, "<<", .RetirarTodosBtn
  Text 228,5,63,7, "www.ffauditoria.com.br", .Text1
  Text 229,13,63,7, "sac@ffauditoria.com.br", .Text1
  Text 206,21,85,7, "(11) 99467-0893 / (11) 3042-8270", .Text1
  GroupBox 0,27,322,133, .GroupBox1
  Text 12,14,62,8, "Auditoria e Consultoria", .Text2
  Text 12,7,65,8, "________________________", .Text3
  Text 12,5,65,8, "Fernandes e Fernandes", .Text4
End Dialog




Option Explicit
Public files()
Dim tempListbox1() As String
Dim tempListBox2() As String
Dim tempListSelect1 As Integer ' var to hold the last SuppValue before hitting the button - need as a global so it is remembered
Dim tempListSelect2 As Integer ' var to hold the last SuppValue before hitting the button - need as a global so it is remembered
Dim vDiretorio As String
Dim ErroAcum As String
Dim sDir As String
Dim sDirAntigo As String
Dim sTrocaOrdem As String
Dim dlgMenu As GerarScriptsDlg
Dim bExitMenu As Boolean
Dim i As Integer

Sub Main	
	vDiretorio = Mid(Client.WorkingDirectory, 1, (Len(Client.WorkingDirectory) - 1))
	sDir = vDiretorio & "\Macros.ILB\"
	Dim button As Integer
	button = Dialog(dlgMenu)
End Sub


Function funMainMenu(ControlID$, Action%, SuppValue%)
	Select Case action
		Case 1	'Iniciando a tela..		
			Call CarregaPrimeiraLista()
			Call OrganizaListas()	
		Case 2	'Efetuando alguma Ação...
			Select Case ControlId$
				'================================================================================================
				Case "AdicionarBtn"
					If UBound(tempListbox1) > 0 Then
						ReDim Preserve tempListbox2(UBound(tempListbox2) + 1)											
						
						If tempListbox1(tempListSelect1) = "* Selecione um" Then
							tempListbox2(UBound(tempListbox2)) = ""
						Else
							tempListbox2(UBound(tempListbox2)) = tempListbox1(tempListSelect1)
						End If
						
						tempListbox1(tempListSelect1 ) = ""
						
						Call sortArray(tempListbox1)
						Call sortArray(tempListbox2)
						Call removeBlanksFromArray(1)
						Call removeBlanksFromArray(2)
					End If
					
					DlgListBoxArray "ListBox2", tempListbox2()
					DlgListBoxArray "ListBox1", tempListbox1()
					tempListSelect1 = 0
					DlgEnable "SubirBtn", 0
					DlgEnable "DescerBtn", 0
					bExitMenu = FALSE
					
					DlgEnable "AdicionarBtn", 0						
					DlgEnable "RetirarBtn", 0
					
					If tempListbox1(1) = "" Then
						DlgEnable "AdicionarTodosBtn", 0					
					End If
					
					DlgEnable "GerarBtn", 1				
					DlgEnable "RetirarTodosBtn", 1
					
				'================================================================================================
				Case "RetirarBtn"
					If UBound(tempListbox2) > 0 Then
						ReDim Preserve tempListbox1(UBound(tempListbox1) + 1)												
						
						If tempListbox2(tempListSelect2) = "* Selecione um" Then
							tempListbox1(UBound(tempListbox1)) = ""
						Else
							tempListbox1(UBound(tempListbox1)) = tempListbox2(tempListSelect2)
						End If
						tempListbox2(tempListSelect2 ) = ""
						
						Call sortArray(tempListbox1)
						Call sortArray(tempListbox2)
						Call removeBlanksFromArray(1)
						Call removeBlanksFromArray(2)
					End If
					
					DlgListBoxArray "ListBox2", tempListbox2()
					DlgListBoxArray "ListBox1", tempListbox1()
					tempListSelect2 = 0
					DlgEnable "SubirBtn", 0
					DlgEnable "DescerBtn", 0					
					bExitMenu = FALSE
					
					DlgEnable "AdicionarBtn", 0						
					DlgEnable "RetirarBtn", 0
					
					If tempListbox2(1) = "" Then
						DlgEnable "RetirarTodosBtn", 0					
						DlgEnable "GerarBtn", 0
					End If
					
					DlgEnable "AdicionarTodosBtn", 1
					
				'================================================================================================
				Case "AdicionarTodosBtn"
					ReDim tempListbox2(UBound(files) + 1)					
					
					For i = 1 To UBound(files) + 1
						tempListbox2(i) = files(i - 1)
					Next i
					
					Call sortArray(tempListbox2)
					Call removeBlanksFromArray(2)
					ReDim tempListbox1(0)					
					Call sortArray(tempListbox1)
					Call removeBlanksFromArray(1)
					
					DlgListBoxArray "ListBox2", tempListbox2()
					DlgListBoxArray "ListBox1", tempListbox1()					
					
					
					DlgListBoxArray "ListBox2", tempListbox2()
					DlgListBoxArray "ListBox1", tempListbox1()
					DlgEnable "SubirBtn", 0
					DlgEnable "DescerBtn", 0					
					bExitMenu = FALSE
					
					DlgEnable "RetirarBtn", 0
					DlgEnable "RetirarTodosBtn", 1
					DlgEnable "AdicionarBtn", 0
					DlgEnable "AdicionarTodosBtn", 0
					DlgEnable "GerarBtn", 1
					
				'================================================================================================
				Case "RetirarTodosBtn"
					ReDim tempListbox1(UBound(files) + 1)					
					
					For i = 1 To UBound(files) + 1
						tempListbox1(i) = files(i - 1)
					Next i
					
					Call sortArray(tempListbox1)
					Call removeBlanksFromArray(1)
					ReDim tempListbox2(0)
					Call sortArray(tempListbox2)
					Call removeBlanksFromArray(2)					
					DlgListBoxArray "ListBox2", tempListbox2()
					DlgListBoxArray "ListBox1", tempListbox1()
					DlgEnable "SubirBtn", 0
					DlgEnable "DescerBtn", 0					
					bExitMenu = FALSE
					
					DlgEnable "RetirarBtn", 0
					DlgEnable "RetirarTodosBtn", 0
					DlgEnable "AdicionarBtn", 0
					DlgEnable "AdicionarTodosBtn", 1
					DlgEnable "GerarBtn", 0
					
				'================================================================================================
				Case "DiretorioBtn"
					Dim s As String					
					sDir = getFolder()
					Call CarregaPrimeiraLista()
					Call OrganizaListas()
					bExitMenu = FALSE		
				
				'================================================================================================
				Case "SubirBtn"
					If Not (tempListbox2(tempListSelect2 - 1) = "* Selecione um") Then
						sTrocaOrdem = tempListbox2(tempListSelect2 - 1)
						tempListbox2(tempListSelect2 - 1) = tempListbox2(tempListSelect2)
						tempListbox2(tempListSelect2) = sTrocaOrdem												
						DlgListBoxArray "ListBox2", tempListbox2				
						bExitMenu = FALSE						
					End If
					
				'================================================================================================
				Case "DescerBtn"
					If Not ( tempListbox2(tempListSelect2) = tempListbox2(UBound(tempListbox2)) ) Then
						sTrocaOrdem = tempListbox2(tempListSelect2 + 1)
						tempListbox2(tempListSelect2 + 1) = tempListbox2(tempListSelect2)
						tempListbox2(tempListSelect2) = sTrocaOrdem
						DlgListBoxArray "ListBox2", tempListbox2											
						bExitMenu = FALSE
					End If
					
				'================================================================================================
				Case "SairBtn"
					bExitMenu = TRUE
					
				'================================================================================================
				Case "ListBox1"					
					tempListSelect1 = SuppValue%
					DlgEnable "SubirBtn", 0
					DlgEnable "DescerBtn", 0
					If tempListSelect1 = 0 Then
						DlgEnable "AdicionarBtn", 0
					Else
						DlgEnable "AdicionarBtn", 1
					End If
					
				'================================================================================================
				Case "ListBox2"
					tempListSelect2 = SuppValue%				
					DlgEnable "SubirBtn", 0
					DlgEnable "DescerBtn", 0
					
					If Not (tempListbox2(tempListSelect2) = "* Selecione um") Then
						DlgEnable "SubirBtn", 1
						DlgEnable "DescerBtn", 1
					End If
					
					If tempListSelect2 = 0 Then
						DlgEnable "RetirarBtn", 0
					Else
						DlgEnable "RetirarBtn", 1
					End If
					
				'================================================================================================
				Case "GerarBtn"
					Call Gerar()
			End Select 'ControlId$			
	End Select 'action
	
	If bExitMenu Then
		funMainMenu = 0
	Else 
		funMainMenu = 1
	End If
End Function

Function CarregaPrimeiraLista
Dim bAchou As Boolean	

	If sDir = "" Then
		sDir = sDirAntigo
	Else
		sDirAntigo = sDir
	End If
	
	ReDim tempListbox1(0)
	If FindFiles(sDir,  "iss",  files()) Then
		For i = 1 To UBound(files) + 1
			ReDim Preserve tempListbox1(UBound(tempListbox1) + 1)				
			tempListbox1(i) = files(i - 1)
		Next i
				
		DlgEnable "AdicionarTodosBtn", 1
		bAchou = True
	Else
		DlgEnable "AdicionarTodosBtn", 0			
		bAchou = False			
	End If
	
	If Not bAchou Then MsgBox "Não foi encontrado Scripts no seguinte diretório:" & Chr(13) & Chr(13) & sDir		
End Function

Function OrganizaListas
	ReDim tempListbox2(0)
	Call sortArray(tempListbox2)
	Call removeBlanksFromArray(2)								
	Call sortArray(tempListbox1)
	Call removeBlanksFromArray(1)
					
	DlgListBoxArray "ListBox1", tempListbox1
	DlgListBoxArray "ListBox2", tempListbox2
	tempListSelect1 = 0
	tempListSelect2 = 0			
	DlgEnable "RetirarBtn", 0
	DlgEnable "RetirarTodosBtn", 0
	DlgEnable "GerarBtn", 0
	DlgEnable "AdicionarBtn", 0	
	DlgEnable "SubirBtn", 0
	DlgEnable "DescerBtn", 0	
End Function

Function Gerar
Dim ScriptErro As String
	For i = 1 To UBound(tempListbox2)
		ScriptErro = tempListbox2(i)
		Client.RunIDEAScript sDir & ScriptErro
	Next i

	Client.WorkingDirectory = vDiretorio
	MsgBox "Processo finalizado!"			
End Function


'****************************************************************************************************
'	Name:		removeBlanksFromArray
'	Description:	Routine to remove blank entries to an array
'****************************************************************************************************
Private Function removeBlanksFromArray(tempType As Integer)
	Dim tempArray() As String
	Dim i, ILoop As Integer
	ReDim tempArray(0)
	
	If tempType = 1 Then
		For ILoop = 0 To UBound(tempListbox1)
			If tempListbox1(ILoop) <> "" Then
				tempArray(UBound(tempArray)) = tempListbox1(ILoop) 
				If ILoop <> UBound(tempListBox1) Then 'don't increment on the last pass
					ReDim preserve tempArray(UBound(tempArray) + 1)
				End If
			End If
		Next ILoop
		'MsgBox UBound(MyArray)
		i = UBound(tempArray)
		Erase tempListbox1
	
		ReDim tempListbox1(i)
		For ILoop = 0 To UBound(tempArray)
			'MsgBox "i " & ILoop & " - " & tempArray(ILoop)
			tempListbox1(ILoop) = tempArray(ILoop) 
		Next ILoop
	Else
		For ILoop = 0 To UBound(tempListbox2)
			If tempListbox2(ILoop) <> "" Then
				tempArray(UBound(tempArray)) = tempListbox2(ILoop) 
				If ILoop <> UBound(tempListBox2) Then 'don't increment on the last pass
					ReDim preserve tempArray(UBound(tempArray) + 1)
				End If
			End If
		Next ILoop
		'MsgBox UBound(MyArray)
		i = UBound(tempArray)
		Erase tempListbox2
	
		ReDim tempListbox2(i)
		For ILoop = 0 To UBound(tempArray)
			'MsgBox "i " & ILoop & " - " & tempArray(ILoop)
			tempListbox2(ILoop) = tempArray(ILoop) 
		Next ILoop

	End If

End Function

'****************************************************************************************************
'	Name:		FindFiles
'	Description:	Routine to find files of a certain type (extension) in a certain directory
'	Accepts:		path - string - path where the files are stored
'			ext - string - extension of files to find
'			files() - array - array to store the list of files
'	Returns:		
'****************************************************************************************************
Private Function FindFiles(path As String,  ext As String,  files())

	Dim ffile As String
	Dim firstbackspace As String
	Dim firstbackspacenum As String
	Dim importname As String
	ffile = Dir$(path & "*." & ext)
	
	If Len(ffile) = 0 Then Exit Function
	
	'-------- Dir ext = *.ext* fixing by checking length
	Do 				
	
		firstbackspace = strReverse (ffile) 
		firstbackspacenum  = InStr(1,firstbackspace, ".")
		importname = Right(ffile, firstbackspacenum - 1)			

		If Len(importname) = Len(ext) Then 
		
			If Not IsNull(ffile) Then
			
				'-------- If one value found return function true and redim array 
				If (FindFiles = False) Then				
					ReDim files(0)
					FindFiles = True
				Else
				
					ReDim Preserve files(UBound(files) + 1) 		
			
				End If
		
			files(UBound(files)) = ffile
		
			Else
				Exit Do
			End If
	
		End If
		
		ffile = Dir        
	
	Loop Until Len(ffile) = 0
        
End Function

'****************************************************************************************************
'	Name:		sortArray
'	Description:	Routine to sort an array
'	Accepts:		A one dimensional array
'	Returns:		Same array sorted 
'****************************************************************************************************
Private Function sortArray(MyArray() As String)
	Dim lLoop, lLoop2 As Integer
	Dim str1, str2 As String

	For lLoop = 1 To UBound(MyArray)
	
		For lLoop2 = lLoop To UBound(MyArray)
		
			If UCase(MyArray(lLoop2)) < UCase(MyArray(lLoop)) Then
			
				str1 = MyArray(lLoop)
				str2 = MyArray(lLoop2)
				MyArray(lLoop) = str2
				MyArray(lLoop2) = str1
			
			End If
		
		Next lLoop2
	
	Next lLoop
	MyArray(0) = "* Selecione um"
End Function


'****************************************************************************************************
'	Name:		getFolder
'	Description:	Routine to obtain the folder
'	Returns:		the folder path
'****************************************************************************************************
Function getFolder() 'this one uses the working directory as the highest level directory
	Dim BrowseFolder As String
	Dim oFolder, oFolderItem
	Dim oPath, oShell, strPath
	Dim Current As Variant 'per Windows documentation this is defined as a variant and not a string
	
	
	Set oShell = CreateObject( "Shell.Application" )
	Set oFolder = oShell.Namespace(17) 'the 17 indicates that we are looking at the virtual folder that contains everything on the local computer
	Current = Client.WorkingDirectory()
	Set oFolder = oShell.BrowseForFolder(0, "Please select the folder where the files are located:", 1, Current)
	
	If (Not oFolder is Nothing) Then
		Set oFolderItem = oFolder.Self
		oPath = oFolderItem.Path
		
		If Right(oPath, 1) <> "\" Then
			oPath = oPath & "\"
		End If
	End If
	
	getFolder = oPath   
End Function
