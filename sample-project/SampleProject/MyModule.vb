' MyModule: The main module of the sample project
'   by imacat <imacat@mail.imacat.idv.tw>, 2017-08-01

Option Explicit

' Main: The main program
Sub Main
    Dim oDialog As Object
    
	DialogLibraries.loadLibrary "SampleProject"
	oDialog = CreateUnoDialog (DialogLibraries.SampleProject.MyDialog)
	' Cancelled
	If oDialog.execute = 0 Then
		Exit Sub
	End If
	
	
	MsgBox GetResString ("DinnerChoice") & oDialog.getControl ("MenuList1").getSelectedItem
End Sub
