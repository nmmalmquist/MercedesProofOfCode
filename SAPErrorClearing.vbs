Dim SapGuiAuto as Object
Dim application as Object
Dim connection as Object
Dim session as Object


SapGuiAuto  = GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session    = connection.Children(0)


Dim RestaredMessagesCount as Integer
RestaredMessagesCount = 0
StringOfErrors = "" 'holds the error message numbers for the ones that cant be restarted



''''Expands each node in heirarchy so we can select items
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").expandNode ("          2")
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").topNode = "          1"
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").expandNode ("          3")
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").topNode = "          1"


'''''''''''''''''''''''''''''''''''This section will cycle through clicking each item and pressing refresh button every iteraion
'Gets the number of messages from query
Dim TopNodeName as String
Dim PoorlyFormattedRowCount as String
Dim ItemNodeName as String
Dim PoorlyFormattedItemMessageCount as String
Dim ItemMessageCount as Integer

'The logic: takes the nope node item and strips apart the next to get the number of error messages
TopNodeName = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").getNodeTextByPath ("1\1")
PoorlyFormattedRowCount = Right(TopNodeName, Len(TopNodeName)- InStr(TopNodeName, "("))
PreRestartedFormattedRowCount = Trim(Left(PoorlyFormattedRowCount,InStr(PoorlyFormattedRowCount,")") -1))

Dim ID as String
ID = ""
 
On Error Resume Next
For i = 4 To 2000 Step 1
'used to get the correct ID, the ID will always be 11 character with spaces leading up to the numbers, so there will be 10 spaces in from of 1 degit numbers, 9 for 2 digit numbers, etc.
	If (i < 10) Then
		ID = "          " & i
	ElseIf (10 <= i AND i < 100) Then
		ID = "         " & i
	Else
		ID = "        " & i
	End If

'clicking each item. Error handling to catch when there is not another item to select (this means we are done)
	On Error Resume Next
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").selectItem (ID,"&Hierarchy")
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").ensureVisibleHorizontalItem (ID,"&Hierarchy")
	System.Threading.Thread.Sleep(500) 

'If an error occured then that means we are done, so leave the loop, remove errors
	If (err.number <> 0) Then
		On Error GoTo 0
		Exit For
	Else:
		On Error GoTo 0
	End If
	
'Press Restart Button	
	session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[0]").pressButton ("RESTART")

'Press save changes button if the pop up appears, else just move on.
	On Error Resume Next
	session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
	If (err.number <> 0) Then
		On Error GoTo 0
	End If
	

'finding the amount of messages in the node item by using the same method as getting total error messages technique above
	ItemNodeName = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").getItemText (ID, "&Hierarchy" )
	PoorlyFormattedItemMessageCount = Right(ItemNodeName, Len(ItemNodeName)- InStr(ItemNodeName, "("))
	ItemMessageCount = Trim(Left(PoorlyFormattedItemMessageCount,InStr(PoorlyFormattedItemMessageCount,")") -1))

	If session.findById("wnd[0]/sbar").Text = "" Then 'if there is a message un the bottom left, then that means its an error message and couldnt be restarted
			
		For j = 1 To ItemMessageCount Step 1
			session.findById("wnd[1]/usr/btnBUTTON_1").press 'press "yes" button that is in the popup after pressing restart
			On Error Resume Next
			session.findById("wnd[1]/tbar[0]/btn[0]").press 'click checkmark button on the log screen that results when there are more than 1 item to restart. This only happens on some, so it is sandwiched with on error resume next
			session.findById("wnd[1]/tbar[0]/btn[0]").press 'cicks checkmark again, but if this one is pressed then that means that there was an error restarting because "it's already in a queue"
			If Err.Number <> 0 Then
				RestartedMessagesCount = RestartedMessagesCount + 1
			Else:
				StringOfErrors = StringOfErrors  & session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").getItemText (ID, "&Hierarchy" ) & ", "
			End If
			On Error GoTo 0
			System.Threading.Thread.Sleep(1000) 
		Next

		

	Else:
		StringOfErrors = StringOfErrors  & session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").getItemText (ID, "&Hierarchy" ) & ", "
	End If

	
Next

'press refresh
session.findById("wnd[0]/tbar[1]/btn[5]").press
System.Threading.Thread.Sleep(15000)

'Get Number of messages after freshing
TopNodeName = session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont[0]/shell/shellcont[0]/shell/shellcont[1]/shell[1]").getNodeTextByPath ("1\1")
PoorlyFormattedRowCount = Right(TopNodeName, Len(TopNodeName)- InStr(TopNodeName, "("))
PostRestartedFormattedRowCount = Trim(Left(PoorlyFormattedRowCount,InStr(PoorlyFormattedRowCount,")") -1))

On Error GoTo 0