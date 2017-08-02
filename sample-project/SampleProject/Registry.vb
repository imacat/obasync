' Registry: Utilities to access to private configuration
'   Taken from TextToColumns, 2017-08-02

Option Explicit

' TODO: Replace SampleProject with your own project name
Const BASE_KEY As String = "/org.openoffice.Office.Addons.SampleProject.AddonConfiguration/"

' GetImageUrl: Returns the image URL for the UNO image controls.
Function GetImageUrl (sName As String) As String
	BasicLibraries.loadLibrary "Tools"
	Dim oRegKey As Object
	
	oRegKey = GetRegistryKeyContent (BASE_KEY & "FileResources/" & sName)
	GetImageUrl = ExpandMacroFieldExpression (oRegKey.Url)
End Function

' GetResString: Returns the localized text string.
Function GetResString (sID As String) As String
	BasicLibraries.loadLibrary "Tools"
	Dim oRegKey As Object
	
	oRegKey = GetRegistryKeyContent (BASE_KEY & "Messages/" & sID)
	GetResString = oRegKey.Text
End Function

' ExpandMacroFieldExpression
Function ExpandMacroFieldExpression (sURL As String) As String
	Dim sTemp As String
	Dim oSM As Object
	Dim oMacroExpander As Object
	
	' Gets the service manager
	oSM = getProcessServiceManager
	' Gets the macro expander
	oMacroExpander = oSM.DefaultContext.getValueByName ( _
		"/singletons/com.sun.star.util.theMacroExpander")
	
	'cut the vnd.sun.star.expand: part
	sTemp = Join (Split (sURL, "vnd.sun.star.expand:"))
	
	'Expand the macrofield expression
	sTemp = oMacroExpander.ExpandMacros (sTemp)
	sTemp = Trim (sTemp)
	ExpandMacroFieldExpression = sTemp
End Function
