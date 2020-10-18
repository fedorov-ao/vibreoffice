
' vibreoffice - Vi Mode for LibreOffice/OpenOffice
'
' The MIT License (MIT)
'
' Copyright (c) 2014 Sean Yeh
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in
' all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
' THE SOFTWARE.

' 2018 axf 
' An attempt to revitalize a useful LibreOffice plugin
' Can be used, but is a work-in-progress so far

' Added Russian keyboard mapping (works in almost all cases)
' Fixed key release processing: releasing of non-char keys is now not consumed
' Fixed char search
' Fixed moving to end of word (e). Moves to the first space after the word.
' Added moving to end of previous word (E). Also moves to the first space after the word. 
' Added search repeating (n and N) 
' Mapped redo to U
' Status bar now displays full vibreoffice state for given frame
' Refactored key translation functions
' Added [c,d][a,i]<char> - will change or delete text fragment surrounded by <char>
' Changed multiplied yank processing: now selects the whole region based on multiplied movement, then yanks. Allows to paste the whole region instead of its last part only.
' Added / and ? searches. String used to search with these is treated as last search and can be used by n and N. 
' Added backspace emulation, called with X
' Replaced string mode id with integer. Should work faster, especially in input.
' Added selection active end swap in visual mode (o).
' Optimized key processing in insert mode.
' Added input logging, log is displayed and cleared by Alt+=
' Added the ability to display the character currently selected by cursor in normal mode
'  If char is printable, shows it as-is, if not - shows its <code>, if empty - shows <>
'  Called by Alt+-
' Added skipping over control and zero-length chars. Downside: cannot join lines easily. Have to go to beggining of next line and press "X"
' Fixed character swapping bug in insert mode by consuming all input and directly using cursor for output
' Fixed dd and cc not working on the last line
' Fixing initialization. Unfortunately, cannot add key handlers and init status bars of all
' currently open windows due to LibreOffice bug (also cannot deinit)
' Thus assuming that vibreoffice is initialized and deinitialized 
' when only one, currently active window is opened
' Added support for multiple windows. Apparently works, but needs cleanup. Vibreoffice is switched into normal mode
' when bringing window in focus
' Fixed s
' Fixed annotation zero-length "char" handling during hjkl movement
' Optimized selection ends swap
' Added polling-based reinitialization on changing windows
' Shift+ESC, as well as "toggle Vbreoffice" menu entry, toggles Vibreoffice

' BUG Annotations break repeated h/l movement
' BUG Anntoations break selection ends swap
' BUG b and e mishadle dashes and periods. Caused by incorrect behaviour of XTextCursor::gotoNext/PreviousWord()
' BUG Searching with f and F behaves incorrectly in visual mode
' TODO Add separate status bar for vibreoffice. Not as easy as it seems.
' TODO Split ProcessMovementKey() into several specialized functions (hjkl movement, word-based movement...).
' Pass number of iterations to these functions to handle repetitions internally.


Option Explicit

' --------
' Globals
' --------
global VIBREOFFICE_STARTED as boolean ' Defaults to False
global VIBREOFFICE_ENABLED as boolean ' Defaults to False

global oXKeyHandler as object
global oListener as object
global oCurrentFrame as object

' Global State
const M_NORMAL = 0
const M_INSERT = 1
const M_VISUAL = 2
const M_DISABLED = 254
const M_BAD = 255
global MODE as integer
global OLD_MODE as integer

global VIEW_CURSOR as object
global TEXT_CURSOR as object
global MULTIPLIER as integer
global LAST_SEARCH as string

global BUSY as boolean

' -----------
' Singletons
' -----------
Sub setCursor
	VIEW_CURSOR = Nothing
	dim oCurrentController
	oCurrentController = getCurrentController()
	If oCurrentController is Nothing Then
		VIEW_CURSOR = Nothing
	Else
	    VIEW_CURSOR = oCurrentController.getViewCursor()
	End If
End Sub

Function getCursor
    getCursor = VIEW_CURSOR
End Function

Sub setTextCursor
    On Error Goto ErrorHandler
    dim oCursor
    oCursor = getCursor()
    dim oText 
    oText = oCursor.getText()
    TEXT_CURSOR = oText.createTextCursorByRange(oCursor)
    Exit Sub
    
ErrorHandler:
    ' Text Cursor does not work in some instances, such as in Annotations
    TEXT_CURSOR = Nothing
End Sub

Function getTextCursor
	setTextCursor() ' temp
    getTextCursor = TEXT_CURSOR
End Function

Function getCurrentController()
	On Error Goto ErrorHandler
	dim oComponent as object : oComponent = thisComponent
	getCurrentController = oComponent.getCurrentController()
	Exit Function
ErrorHandler:
	getCurrentController = Nothing
End Function	

private function printString(oCursor, s)
	dim l : l = len(s)
	oCursor.setString(s)
	oCursor.goRight(l, False)
end function

' -----------------
' Helper Functions
' -----------------
' Returns mode name
Function getModeName(m)
	dim sModeName as string
	Select Case m
		Case M_NORMAL:
			sModeName = "NORMAL"
		Case M_INSERT:
			sModeName = "INSERT"
		Case M_VISUAL:
			sModeName = "VISUAL"
		Case M_BAD:
			sModeName = "BAD"
		Case M_DISABLED:
			sModeName = "DISABLED"
		Case Else:
			sModeName = "BAD"								
	End Select
	getModeName = sModeName
End Function

' Returns key by non-zero code
Function getLatinKeyCharByCode(oEvent)
    dim keyChar
    keyChar = asc(0)
    If (oEvent.modifiers and 1) = 0 Then
        Select Case oEvent.keyCode
            case 1311: 
                keyChar = "`"
            case 257: 
                keyChar = "1"
            case 258: 
                keyChar = "2"
            case 259: 
                keyChar = "3"
            case 260: 
                keyChar = "4"
            case 261: 
                keyChar = "5"
            case 262:
                keyChar = "6"
            case 263:
                keyChar = "7"
            case 264:
                keyChar = "8"
            case 265:
                keyChar = "9"
            case 256:
                keyChar = "0"
            case 1288:
                keyChar = "-"
            case 1295:
                keyChar = "="
            case 0:
                keyChar = "\"
            case 528:
                keyChar = "q"
            case 534:
                keyChar = "w"
            case 516:
                keyChar = "e"
            case 529:
                keyChar = "r"
            case 531:
                keyChar = "t"
            case 536:
                keyChar = "y"
            case 532:
                keyChar = "u"
            case 520:
                keyChar = "i"
            case 526:
                keyChar = "o"
            case 527:
                keyChar = "p"
            case 1315:
                keyChar = "["
            case 1316:
                keyChar = "]"
            case 512:
                keyChar = "a"
            case 530:
                keyChar = "s"
            case 515:
                keyChar = "d"
            case 517:
                keyChar = "f"
            case 518:
                keyChar = "g"
            case 519:
                keyChar = "h"
            case 521:
                keyChar = "j"
            case 522:
                keyChar = "k"
            case 523:
                keyChar = "l"
            case 1317:
                keyChar = ";"
            case 1318:
                keyChar = "'"
            case 537:
                keyChar = "z"
            case 535:
                keyChar = "x"
            case 514:
                keyChar = "c"
            case 533:
                keyChar = "v"
            case 513:
                keyChar = "b"
            case 525:
                keyChar = "n"
            case 524:
                keyChar = "m"
            case 1292:
                keyChar = ","
            case 1291:
                keyChar = "."
            case 1290:
                keyChar = "/"
            case 1284:
                keyChar = " "
        End Select
    Else
        Select Case oEvent.keyCode
            case 1311:
                keyChar = "~"
            case 257:
                keyChar = "!"
            case 258:
                keyChar = "@"
            case 259:
                keyChar = "#"
            case 260:
                keyChar = "$"
            case 261:
                keyChar = "%"
            case 262:
                keyChar = "^"
            case 263:
                keyChar = "&"
            case 264:
                keyChar = "*"
            case 265:
                keyChar = "("
            case 256:
                keyChar = ")"
            case 1288:
                keyChar = "_"
            case 1295:
                keyChar = "+"
            case 528:
                keyChar = "Q"
            case 534:
                keyChar = "W"
            case 516:
                keyChar = "E"
            case 529:
                keyChar = "R"
            case 531:
                keyChar = "T"
            case 536:
                keyChar = "Y"
            case 532:
                keyChar = "U"
            case 520:
                keyChar = "I"
            case 526:
                keyChar = "O"
            case 527:
                keyChar = "P"
            case 1315:
                keyChar = "{"
            case 1316:
                keyChar = "}"
            case 512:
                keyChar = "A"
            case 530:
                keyChar = "S"
            case 515:
                keyChar = "D"
            case 517:
                keyChar = "F"
            case 518:
                keyChar = "G"
            case 519:
                keyChar = "H"
            case 521:
                keyChar = "J"
            case 522:
                keyChar = "K"
            case 523:
                keyChar = "L"
            case 1317:
                keyChar = ":"
            case 1318:
                keyChar = chr(34)
            case 537:
                keyChar = "Z"
            case 535:
                keyChar = "X"
            case 514:
                keyChar = "C"
            case 533:
                keyChar = "V"
            case 513:
                keyChar = "B"
            case 525:
                keyChar = "N"
            case 524:
                keyChar = "M"
            case 1292:
                keyChar = "<"
            case 1291:
                keyChar = ">"
            case 1290:
                keyChar = "?"
            case 1284:
                keyChar = " "
        End Select
    End If
    getLatinKeyCharByCode = keyChar
End Function

Function getLatinKeyCharByRus(oEvent)
    dim keyChar
    keyChar = asc(0)
    If (oEvent.modifiers and 1) = 0 Then
        Select Case oEvent.keyChar
            case "?":
                keyChar = "["
            case "?":
                keyChar = "]"
            case "?":
                keyChar = ";"    
            case "?":
                keyChar = "'"
            case "?":
                keyChar = ","
            case "?":
                keyChar = "."
            case ".":
                keyChar = "."                
        End Select
    Else
        Select Case oEvent.keyChar
            case "?":
                keyChar = "{"
            case "?":
                keyChar = "}"
            case "?":
                keyChar = ":"    
            case "?":
                keyChar = chr(34)
            case "?":
                keyChar = "<"
            case "?":
                keyChar = ">"
        End Select
    End If
    getLatinKeyCharByRus = keyChar    
End Function


Function getLatinKey(oEvent)
    dim keyChar
    keyChar = asc(0)
    If oEvent.keyCode <> 0 Then
        keyChar = getLatinKeyCharByCode(oEvent)
    Else
        keyChar = getLatinKeyCharByRus(oEvent)
    End If
    getLatinKey = keyChar
End Function

Function isControl(c)
	if len(c) = 0 then
		isControl = True
	else
		dim ac as integer : ac = asc(c)
		isControl = ((ac >= 0 and ac <= 31) or (ac = 127))
	end if
End Function

Function isPrintable(c)
	isPrintable = not isControl(c)
End Function

Sub restoreStatus 'restore original statusbar
	On Error Goto ErrorHandler
    dim oCurrentContorller : oCurrentContorller = getCurrentController()
   	dim oFrame : oFrame = oCurrentContorller.Frame
   	dim oLayout : oLayout = oFrame.LayoutManager
  	oLayout.destroyElement("private:resource/statusbar/statusbar")
   	oLayout.createElement("private:resource/statusbar/statusbar")
   	Exit Sub
ErrorHandler:
	MsgBox("restoreStatus() failed!")
End Sub

' Unfortunately, does not work as expected
' Statusbar of the currently active window is restored, because background windows
' return the controller of the currently active window on getCurrentController
' Thus the statusbar of active window is restored several times, and status bars
' of background windows are not restored at all
Sub restoreStatusOfModels()
    dim vComponents
    vComponents = StarDesktop.getComponents()
    If vComponents.hasElements() Then
    	dim vEnumeration
    	vEnumeration = vComponents.createEnumeration()
    	Do While vEnumeration.hasMoreElements()
    		dim vComponent
    		vComponent = vEnumeration.nextElement()
    		If HasUnoInterfaces(vComponent, "com.sun.star.text.XTextDocument") Then
				dim oController
				
    			oController = vComponent.getCurrentController()
	    		static oOldController as object
	    		If EqualUnoObjects(oController, oOldController) Then
	    		'	MsgBox("controllers are the same")
	    		End If
	    		oOldController = oController
	    		
    			If not (oController is Nothing) Then
			    	dim oFrame
			    	oFrame = oController.getFrame()
			    	dim oLayout
			    	oLayout = oFrame.LayoutManager
			    	oLayout.destroyElement("private:resource/statusbar/statusbar")
				    oLayout.createElement("private:resource/statusbar/statusbar")
			   	End If
    		End If
    	Loop	
    End If
End Sub

Sub setRawStatus(rawText)
	dim oCurrentController as object
	oCurrentController = getCurrentController()
	If not (oCurrentController is Nothing) Then
	    oCurrentcontroller.StatusIndicator.Start(rawText, 0)
	End If
End Sub

Sub setStatus()
    setRawStatus(getModeName(MODE) & " | " & getMultiplier() & " | special: " & getSpecial() & " | " & "modifier: " & getMovementModifier())
End Sub

Sub setMode(m)
    MODE = m
    setStatus()
End Sub

Function gotoMode(sMode)
    Select Case sMode
        Case M_NORMAL, M_DISABLED:
            setMode(sMode)
            setMovementModifier("")
        Case M_INSERT, M_VISUAL:
            setMode(sMode)
            ' Deselect TextCursor
			dim oTextCursor
			oTextCursor = getTextCursor()
			If not (oTextCursor is Nothing) Then
            	oTextCursor.gotoRange(oTextCursor.getStart(), False)
            	' Show TextCursor selection
            	getCurrentController().Select(oTextCursor)
            End If                        
		Case Else:
			' Should not happen
			setMode(M_BAD)
    End Select
End Function

Sub cursorPreReset(oTextCursor)
    oTextCursor.gotoRange(oTextCursor.getStart(), False)
    oTextCursor.goRight(1, False)
	oTextCursor.goLeft(1, True)
End Sub

Sub cursorReset(oTextCursor)
	cursorPreReset(oTextCursor)	
    getCurrentController().Select(oTextCursor) ' needed
End Sub

Sub searchAndSet(oTextCursor, sText, bIsBackwards)
    dim oSearchDesc as object
    oSearchDesc = thisComponent.createSearchDescriptor()
    oSearchDesc.setSearchString(sText)
    oSearchDesc.SearchCaseSensitive = True
    oSearchDesc.SearchBackwards = bIsBackwards
    dim oStartRange	    
    If Not bIsBackwards Then
        oStartRange = oTextCursor.getEnd()
    Else
        oStartRange = oTextCursor.getStart()
    End If
    dim oFoundRange	    
    oFoundRange = thisComponent.findNext(oStartRange, oSearchDesc)	    
	If not (oFoundRange is Nothing) Then
		oTextCursor.gotoRange(oFoundRange, False)
		getCurrentController().Select(oTextCursor)
		setMode(M_VISUAL)
	End If
End Sub

' swaps cursor start and end
' have to resort to the crude string-based algorithm because of annotation chars
' BUG minor bug: on swap will exclude annotation chars present on either end of selection
Sub swapCursorEnds(oTextCursor as object)
	dim s, oldLen
	s = oTextCursor.getString()
	oldLen = len(s)
	If oldLen = 0 Then Exit Sub
	oTextCursor.goRight(1, True)
	dim newLen : newLen = len(oTextCursor.getString())
	If newLen > oldLen Then oTextCursor.goLeft(1, False)
	dim pureLen : pureLen = newLen - count(s, chr(10))
	dim t as string
	If newLen >= oldLen Then
		oTextCursor.collapseToEnd()
		oTextCursor.goLeft(pureLen-1, True)
        do while oTextCursor.getString() <> s
        	t = oTextCursor.getString()
			oTextCursor.goLeft(1, True)
		Loop
	Else
		oTextCursor.collapseToStart()
		oTextCursor.goLeft(1, False) 'collapsing to start does not include starting character for some reason
        oTextCursor.goRight(pureLen, True)
        do while oTextCursor.getString() <> s
        	t = oTextCursor.getString()
			oTextCursor.goRight(1, True)
		Loop
	End If
End Sub

Function samePos(oPos1, oPos2)
    samePos = oPos1.X() = oPos2.X() And oPos1.Y() = oPos2.Y()
End Function

Function genString(sChar, iLen)
    dim sResult, i
    sResult = ""
    For i = 1 To iLen
        sResult = sResult & sChar
    Next i
    genString = sResult
End Function

' Counts number of character c in string s
Function count(s as string, c)
	dim cnt
	cnt = 0
	dim length
	length = len(s)
	dim i
	For i = 1 to length
		If asc(Mid(s,i,1)) = asc(c) Then cnt = cnt + 1
	Next
	count = cnt
End Function

' Yanks selection to system clipboard.
' If bDelete is true, will delete selection.
Sub yankSelection(bDelete)
    dim dispatcher As Object
    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(getCurrentController().Frame, ".uno:Copy", "", 0, Array())

    If bDelete Then
        getTextCursor().setString("")
    End If
End Sub


Sub pasteSelection()
    dim oTextCursor, dispatcher As Object

    ' Deselect if in NORMAL mode to avoid overwriting the character underneath
    ' the cursor
    If MODE = M_NORMAL Then
        oTextCursor = getTextCursor()
        oTextCursor.gotoRange(oTextCursor.getStart(), False)
        getCurrentController().Select(oTextCursor)
    End If

    dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")
    dispatcher.executeDispatch(getCurrentController().Frame(), ".uno:Paste", "", 0, Array())
End Sub


' -----------------------------------
' Special Mode (for chained commands)
' -----------------------------------
global SPECIAL_MODE As string
global SPECIAL_COUNT As integer

Sub setSpecial(specialName)
    SPECIAL_MODE = specialName

    If specialName = "" Then
        SPECIAL_COUNT = 0
    Else
        SPECIAL_COUNT = 2
    End If
End Sub

Function getSpecial()
    getSpecial = SPECIAL_MODE
End Function

Sub delaySpecialReset()
    SPECIAL_COUNT = SPECIAL_COUNT + 1
End Sub

Sub resetSpecial(Optional bForce)
    If IsMissing(bForce) Then bForce = False

    SPECIAL_COUNT = SPECIAL_COUNT - 1
    If SPECIAL_COUNT <= 0 Or bForce Then
        setSpecial("")
    End If
End Sub


' -----------------
' Movement Modifier
' -----------------
'f,i,a
global MOVEMENT_MODIFIER As string

Sub setMovementModifier(modifierName)
    MOVEMENT_MODIFIER = modifierName
End Sub

Function getMovementModifier()
    getMovementModifier = MOVEMENT_MODIFIER
End Function


' --------------------
' Multiplier functions
' --------------------
Sub _setMultiplier(n as integer)
    MULTIPLIER = n
End Sub

Sub resetMultiplier()
    _setMultiplier(0)
End Sub

Sub addToMultiplier(n as integer)
    dim sMultiplierStr as String
    dim iMultiplierInt as integer

    ' Max multiplier: 10000 (stop accepting additions after 1000)
    If MULTIPLIER <= 1000 then
        sMultiplierStr = CStr(MULTIPLIER) & CStr(n)
        _setMultiplier(CInt(sMultiplierStr))
    End If
End Sub

' Should only be used if you need the raw value
Function getRawMultiplier()
    getRawMultiplier = MULTIPLIER
End Function

' Same as getRawMultiplier, but defaults to 1 if it is unset (0)
Function getMultiplier()
    If MULTIPLIER = 0 Then
        getMultiplier = 1
    Else
        getMultiplier = MULTIPLIER
    End If
End Function


' -------------
' Key Handling
' -------------
' Buggy due to the same reason as restoreStatusOfModels() is
Sub addKeyHandlerToModels()
	If oXKeyHandler is nothing Then
		Exit Sub
	End If
	
    dim vComponents
    vComponents = StarDesktop.getComponents()
    If vComponents.hasElements() Then
    	dim vEnumeration
    	vEnumeration = vComponents.createEnumeration()
    	Do While vEnumeration.hasMoreElements()
    		dim vComponent
    		vComponent = vEnumeration.nextElement()
    		If HasUnoInterfaces(vComponent, "com.sun.star.text.XTextDocument") Then
				dim oController
    			oController = vComponent.getCurrentController()
    			If not (oController is Nothing) Then
			    	oController().addKeyHandler(oXKeyHandler)
			   	End If
    		End If
    	Loop	
    End If
End Sub


' Buggy due to the same reason as restoreStatusOfModels() is
Sub removeKeyHandlerFromModels()
    dim vComponents
    vComponents = StarDesktop.getComponents()
    If vComponents.hasElements() Then
    	dim vEnumeration
    	vEnumeration = vComponents.createEnumeration()
    	Do While vEnumeration.hasMoreElements()
    		dim vComponent
    		vComponent = vEnumeration.nextElement()
    		If HasUnoInterfaces(vComponent, "com.sun.star.text.XTextDocument") Then
				dim oController
    			oController = vComponent.getCurrentController()
    			If not (oController is Nothing) Then
			    	oController().removeKeyHandler(oXKeyHandler)
			    End If
    		End If
    	Loop	
    End If
End Sub


Sub sStartXKeyHandler
	sStopXKeyHandler()
    oXKeyHandler = CreateUnoListener("KeyHandler_", "com.sun.star.awt.XKeyHandler")
End Sub

Sub sStopXKeyHandler
End Sub

Sub KeyHandler_Disposing(oEvent)
End Sub


' --------------------
' Main Key Processing
' --------------------
function KeyHandler_KeyPressed(oEvent) as boolean
  ' Exit if plugin is not enabled
  If MODE = M_DISABLED Then
    KeyHandler_KeyPressed = False
    Exit Function
  End If

  ' Block all inputs by default
  dim bConsumeInput : bConsumeInput = True

  ' If this function is invoked while previous event is still being processed, just discard the new event
  If BUSY = True Then
    KeyHandler_KeyPressed = True
    Exit Function
  End If
  BUSY = True

  ' Have to resort to polling because subscribing to theGlobalEventBroadcaster causes crashes
  dim oFrame : oFrame = StarDesktop.getCurrentFrame()
  If not EqualUnoObjects(oFrame, oCurrentFrame) Then
    reinitVibreOffice()
    oCurrentFrame = oFrame
  End If

  ' --------------------------
  ' Process global shortcuts, exit if matched (like ESC)
  If ProcessGlobalKey(oEvent) Then
    ' Pass
  ' If INSERT mode, allow all inputs
  ElseIf getTextCursor() is Nothing Then
    bConsumeInput = False
  ElseIf MODE = M_INSERT Then
    bConsumeInput = False
  Else
    dim bIsMultiplier, bIsModified, bIsControl, bIsSpecial
    bIsMultiplier = False ' reset multiplier by default
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)
    bIsSpecial = getSpecial() <> ""        	
    ' If Change Mode
    ' ElseIf MODE = M_NORMAL And Not bIsSpecial And getMovementModifier() = "" And ProcessModeKey(oEvent) Then
    If ProcessModeKey(oEvent) Then
        ' Pass

    ' Replace Key
    ElseIf getSpecial() = "r" And Not bIsModified Then
        dim iLen
        iLen = Len(getCursor().getString())
        getCursor().setString(genString(oEvent.KeyChar, iLen))

    ' Multiplier Key
    ElseIf ProcessNumberKey(oEvent) Then
        bIsMultiplier = True
        delaySpecialReset()

    ' Normal Key
    ElseIf ProcessNormalKey(getLatinKey(oEvent), oEvent.Modifiers, oEvent) Then
        ' Pass

    ' If is modified but doesn't match a normal command, allow input
    '   (Useful for built-in shortcuts like Ctrl+a, Ctrl+s, Ctrl+w)
    ElseIf bIsModified Then
        ' Ctrl+a (select all) sets mode to VISUAL
        If bIsControl And getLatinKey(oEvent) = "a" Then
            gotoMode(M_VISUAL)
        End If
        bConsumeInput = False

    ' Movement modifier here?
    ElseIf ProcessMovementModifierKey(getLatinKey(oEvent)) Then
        delaySpecialReset()

    ' If standard movement key (in VISUAL mode) like arrow keys, home, end
    ElseIf MODE = M_VISUAL And ProcessStandardMovementKey(oEvent) Then
        ' Pass

    ' If bIsSpecial but nothing matched, return to normal mode
    ElseIf bIsSpecial Then
        gotoMode(M_NORMAL)

    ' Allow non-letter keys if unmatched
    ' TODO Use getLatinKey()
    ElseIf asc(oEvent.KeyChar) = 0 Then
        bConsumeInput = False
    End If
    ' --------------------------

    ' Reset Special
    resetSpecial()

    ' Reset multiplier if last input was not number and not in special mode
    If not bIsMultiplier and getSpecial() = "" and getMovementModifier() = "" Then
        resetMultiplier()
    End If
    setStatus(getMultiplier())
	End If

  KeyHandler_KeyPressed = bConsumeInput
  BUSY = False
End Function

Function KeyHandler_KeyReleased(oEvent) As boolean
  ' Exit if plugin is not enabled
  If MODE = M_DISABLED Then
    KeyHandler_KeyReleased = False
    Exit Function
  End If
  
  If asc(oEvent.KeyChar) = 0 Then
    KeyHandler_KeyReleased = False
  Else
    dim iModifiers : iModifiers = oEvent.modifiers
    dim iKeyCode : iKeyCode = oEvent.keyCode
    ' Allow Ctrl+c and Ctrl-v, so don't change cursor
    If (iKeyCode = 514 or iKeyCode = 533) And (iModifiers = 2 Or iModifiers = 8) Then
    ' Needed to make cursor always select 1 character in NORMAL mode
    ' Constrict to movement keys only?
    ElseIf MODE = M_NORMAL Then
      ' Show terminal-like cursor
      dim oTextCursor
      oTextCursor = getTextCursor()
      If not (oTextCursor Is Nothing) Then
         ' Do nothing        
         cursorReset(oTextCursor)
      End If
    End If

    KeyHandler_KeyReleased = (MODE = M_NORMAL) 'cancel KeyReleased
  End If
End Function


' ----------------
' Processing Keys
' ----------------
Function ProcessGlobalKey(oEvent)
    dim bMatched, bIsControl
    bMatched = False
    bIsControl = (oEvent.Modifiers = 2) or (oEvent.Modifiers = 8)

    ' PRESSED ESCAPE (or ctrl+[)
    If oEvent.KeyCode = 1281 Or (oEvent.KeyCode = 1315 And bIsControl) Then
    	If getTextCursor() is Nothing Then
	    	bMatched = False
    	Else
        	' Move cursor back if was in INSERT (but stay on same line)
        	If MODE <> M_NORMAL And Not getCursor().isAtStartOfLine() Then
            	getCursor().goLeft(1, False)
        	End If
        	bMatched = True
        End If

        resetSpecial(True)
        gotoMode(M_NORMAL)
    Else
        bMatched = False
    End If
    ProcessGlobalKey = bMatched
End Function


Function ProcessStandardMovementKey(oEvent)
    dim c, bMatched
    c = oEvent.KeyCode

    bMatched = True

    If MODE <> M_VISUAL Then
        bMatched = False
        'Pass
    ElseIf c = 1024 Then
        ProcessMovementKey("j", True)
    ElseIf c = 1025 Then
        ProcessMovementKey("k", True)
    ElseIf c = 1026 Then
        ProcessMovementKey("h", True)
    ElseIf c = 1027 Then
        ProcessMovementKey("l", True)
    ElseIf c = 1028 Then
        ProcessMovementKey("^", True)
    ElseIf c = 1029 Then
        ProcessMovementKey("$", True)
    Else
        bMatched = False
    End If

    ProcessStandardMovementKey = bMatched
End Function


Function ProcessNumberKey(oEvent)
    dim c
    c = CStr(oEvent.KeyChar)

	' Don't treat number keys as multiplier-related 
	' if we are in modified movement mode (like f,t)
	' Otherwise will not be able to search for numbers with f or t at all
    If getMovementModifier() = "" and c >= "0" and c <= "9" Then
        addToMultiplier(CInt(c))
        ProcessNumberKey = True
    Else
        ProcessNumberKey = False
    End If
End Function


Function ProcessModeKey(oEvent)
    dim bIsModified
    bIsModified = oEvent.Modifiers > 1 ' If Ctrl or Alt is held down. (Shift=1)
    ' Don't change modes in these circumstances
    If MODE <> M_NORMAL Or bIsModified Or getSpecial <> "" Or getMovementModifier() <> "" Then
        ProcessModeKey = False
        Exit Function
    End If

    ' Mode matching
    dim bMatched
    dim keyChar
    bMatched = True
    keyChar = getLatinKey(oEvent)
    Select Case keyChar
        ' Insert modes
        Case "i", "a", "I", "A", "o", "O":
            If keyChar = "a" Then getCursor().goRight(1, False)
            If keyChar = "I" Then ProcessMovementKey("^")
            If KeyChar = "A" Then ProcessMovementKey("$")

            If KeyChar = "o" Then
                ProcessMovementKey("$")
                ProcessMovementKey("l")
                getCursor().setString(chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    getCursor().setString(chr(13) & chr(13))
                    ProcessMovementKey("l")
                End If
            End If

            If KeyChar = "O" Then
                ProcessMovementKey("^")
                getCursor().setString(chr(13))
                If Not getCursor().isAtStartOfLine() Then
                    ProcessMovementKey("h")
                    getCursor().setString(chr(13))
                    ProcessMovementKey("l")
                End If
            End If

            gotoMode(M_INSERT)
        Case "v":
            gotoMode(M_VISUAL)
        Case Else:
            bMatched = False
    End Select
    ProcessModeKey = bMatched
End Function


Function ProcessNormalKey(keyChar, modifiers, optional oEvent)
    dim i, bMatched, bMatchedMovement, bIsVisual, iIterations, bIsControl, sSpecial
    bIsControl = (modifiers = 2) or (modifiers = 8)
    bIsVisual = (MODE = M_VISUAL) ' is this hardcoding bad? what about visual block?

	If bIsVisual and keyChar = "o" Then
		dim oTextCursor
		oTextCursor = getTextCursor()
		swapCursorEnds(oTextCursor)
		getCurrentController().select(oTextCursor)
        ProcessNormalKey = True
        Exit Function		
	End If
    ' ----------------------
    ' 1. Check Movement Key
    ' ----------------------
    iIterations = getMultiplier()
    sSpecial = getSpecial()
    bMatched = False
    bMatchedMovement = False
    ' starting from 0 adds one extra, unneeded movement
    ' FIXME axf Currently, say, <num>f<char> and d<num>f<char> commands differ if cursor is on <char>
    For i = 1 To iIterations 
        ' Movement Key
        ' axf Passing oEvent to make actual key char available for search
        bMatchedMovement = ProcessMovementKey(keyChar, bIsVisual, modifiers, oEvent)
        bMatched = bMatched or bMatchedMovement
    Next i

    ' If Special: d/c + movement
    If bMatched And (sSpecial = "d" Or sSpecial = "c" Or sSpecial = "y") Then
        yankSelection((sSpecial <> "y"))
    End If    

    ' Reset Movement Modifier
    setMovementModifier("")

    ' Exit already if movement key was matched
    If bMatched Then
        ' If Special: d/c : change mode
        If getSpecial() = "d" Or getSpecial() = "y" Then gotoMode(M_NORMAL)
        If getSpecial() = "c" Then gotoMode(M_INSERT)

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 2. Undo/Redo
    ' --------------------
    If keyChar = "u" Or keyChar = "U" Then
        dim mode
        mode = 0
        If keyChar = "u" Then
            mode = 1
        End If

        For i = 1 To iIterations
            Undo(mode)
        Next i

        ProcessNormalKey = True
        Exit Function
    End If


    ' --------------------
    ' 3. Paste
    '   Note: in vim, paste will result in cursor being over the last character
    '   of the pasted content. Here, the cursor will be the next character
    '   after that. Fix?
    ' --------------------
    If keyChar = "p" or keyChar = "P" Then
        ' Move cursor right if "p" to paste after cursor
        If keyChar = "p" Then
            ProcessMovementKey("l", False)
        End If

        For i = 1 To iIterations
            pasteSelection()
        Next i

        ProcessNormalKey = True
        Exit Function
    End If
    
    ' search
    ' HACK Remapping to make compatible with Russian keyboard layout
    If keyChar = "." Then 
    	keyChar = "/" 
    ElseIf keyChar = ">" Then
    	keyChar = "?"
    End If
    
    If keyChar = "/" or keyChar = "?" Then
    	dim sDir
    	If keyChar = "/" Then
    		sDir = "forward"
    	Else
    		sDir = "backward"
    	End If
    	sDir = "Search " & sDir
    	dim sInput
    	sInput = InputBox(sDir, sDir)
    	If sInput <> "" Then
    		dim bIsBackwards
	 	    bIsBackwards = (keyChar = "?")
		    searchAndSet(getTextCursor(), sInput, bIsBackwards)
			LAST_SEARCH = sInput
	   		ProcessNormalKey = True
	        Exit Function	
	     End If	
    End If

    ' --------------------
    ' 4. Check Special/Delete Key
    ' --------------------

    ' There are no special/delete keys with modifier keys, so exit early
    If modifiers > 1 Then
        ProcessNormalKey = False
        Exit Function
    End If

    ' Only 'x' or Special (dd, cc) can be done more than once
    If keyChar <> "x" and getSpecial() = "" Then
        iIterations = 1
    End If
    For i = 1 To iIterations
        dim bMatchedSpecial

        ' Special/Delete Key
        bMatchedSpecial = ProcessSpecialKey(keyChar)

        bMatched = bMatched or bMatchedSpecial
    Next i


    ProcessNormalKey = bMatched
End Function


' Function for both undo and redo
Sub Undo(bUndo)
    On Error Goto ErrorHandler

    If bUndo Then
        thisComponent.getUndoManager().undo()
    Else
        thisComponent.getUndoManager().redo()
    End If
    Exit Sub

    ' Ignore errors from no more undos/redos in stack
ErrorHandler:
    Resume Next
End Sub


Function ProcessSpecialKey(keyChar)
    dim oCursor, oTextCursor, bMatched, bIsSpecial, bIsDelete
    bMatched = True
    bIsSpecial = getSpecial() <> ""


    If keyChar = "d" Or keyChar = "c" Or keyChar = "s" Or keyChar = "y" Then
        bIsDelete = (keyChar <> "y")

        ' Special Cases: 'dd' and 'cc'
        If bIsSpecial Then
            dim bIsSpecialCase
            bIsSpecialCase = (keyChar = "d" And getSpecial() = "d") Or (keyChar = "c" And getSpecial() = "c")

            If bIsSpecialCase Then
            	' A bit hacky, but works
                oCursor = getCursor()
                oCursor.gotoStartOfLine(False)
                oCursor.gotoEndOfLine(True)                

                oTextCursor = getTextCursor()
                oTextCursor.goRight(1, True)
                getCurrentController().Select(oTextCursor)
                yankSelection(bIsDelete)
            Else
                bMatched = False
            End If

            ' Go to INSERT mode after 'cc', otherwise NORMAL
            If bIsSpecialCase And keyChar = "c" Then
                gotoMode(M_INSERT)
            Else
                gotoMode(M_NORMAL)
            End If


        ' visual mode: delete selection
        ElseIf MODE = M_VISUAL Then
            oTextCursor = getTextCursor()
            getCurrentController().Select(oTextCursor)

            yankSelection(bIsDelete)

            If keyChar = "c" Or keyChar = "s" Then gotoMode(M_INSERT)
            If keyChar = "d" Or keyChar = "y" Then gotoMode(M_NORMAL)


        ' Enter Special mode: 'd', 'c', or 'y' ('s' => 'cl')
        ElseIf MODE = M_NORMAL Then

            ' 's' => 'cl'
            If keyChar = "s" Then
                setSpecial("c")
                gotoMode(M_VISUAL)
                ProcessNormalKey("l", 0, Nothing)
            Else
                setSpecial(keyChar)
                gotoMode(M_VISUAL)
            End If
        End If

    ' If is 'r' for replace
    ElseIf keyChar = "r" Then
        setSpecial("r")

    ' Otherwise, ignore if bIsSpecial
    ElseIf bIsSpecial Then
        bMatched = False


    ElseIf keyChar = "x" or keyChar = "X" Then
        oTextCursor = getTextCursor()
        If keyChar = "X" Then
        	oTextCursor.goLeft(1, False)
        	oTextCursor.goRight(1, True)        	
        End If
        getCurrentController().Select(oTextCursor)
        yankSelection(True)

        ' Reset Cursor
        cursorReset(oTextCursor)

        ' Goto NORMAL mode (in the case of VISUAL mode)
        gotoMode(M_NORMAL)

    ElseIf keyChar = "D" Or keyChar = "C" Then
        If MODE = M_VISUAL Then
            ProcessMovementKey("^", False)
            ProcessMovementKey("$", True)
            ProcessMovementKey("l", True)
        Else
            ' Deselect
            oTextCursor = getTextCursor()
            oTextCursor.gotoRange(oTextCursor.getStart(), False)
            getCurrentController().Select(oTextCursor)
            ProcessMovementKey("$", True)
        End If

        yankSelection(True)

        If keyChar = "D" Then
            gotoMode(M_NORMAL)
        ElseIf keyChar = "C" Then
            gotoMode(M_INSERT)
        End IF

    ' S only valid in NORMAL mode
    ElseIf keyChar = "S" And MODE = M_NORMAL Then
        ProcessMovementKey("^", False)
        ProcessMovementKey("$", True)
        yankSelection(True)
        gotoMode(M_INSERT)

    Else
        bMatched = False
    End If

    ProcessSpecialKey = bMatched
End Function


Function ProcessMovementModifierKey(keyChar)
    dim bMatched

    bMatched = True
    Select Case keyChar
        Case "f", "t", "F", "T", "i", "a":
            setMovementModifier(keyChar)
        Case Else:
            bMatched = False
    End Select

    ProcessMovementModifierKey = bMatched
End Function


Function ProcessSearchKey(oTextCursor, searchType, keyChar, bExpand)
    '-----------
    ' Searching
    '-----------
    dim bMatched, oSearchDesc, oFoundRange, bIsBackwards, oStartRange
    bMatched = True
    bIsBackwards = (searchType = "F" Or searchType = "T")

    If Not bIsBackwards Then
        ' VISUAL mode will goRight AFTER the selection
        If MODE <> M_VISUAL Then
            ' Start searching from next character
            oTextCursor.goRight(1, bExpand)
        End If

        oStartRange = oTextCursor.getEnd()
        ' Go back one
        oTextCursor.goLeft(1, bExpand)
    Else
        oStartRange = oTextCursor.getStart()
    End If

    oSearchDesc = thisComponent.createSearchDescriptor()
    oSearchDesc.setSearchString(keyChar)
    oSearchDesc.SearchCaseSensitive = True
    oSearchDesc.SearchBackwards = bIsBackwards

    oFoundRange = thisComponent.findNext( oStartRange, oSearchDesc )

    If not IsNull(oFoundRange) Then
        dim oText, foundPos, curPos, bSearching
        oText = oTextCursor.getText()
        foundPos = oFoundRange.getStart()

        ' Unfortunately, we must go go to this "found" position one character at
        ' a time because I have yet to find a way to consistently move the
        ' Start range of the text cursor and leave the End range intact.
        If bIsBackwards Then
            curPos = oTextCursor.getEnd()
        Else
            curPos = oTextCursor.getStart()
        End If
        do until oText.compareRegionStarts(foundPos, curPos) = 0
            If bIsBackwards Then
                bSearching = oTextCursor.goLeft(1, bExpand)
                curPos = oTextCursor.getStart()
            Else
                bSearching = oTextCursor.goRight(1, bExpand)
                curPos = oTextCursor.getEnd()
            End If

            ' Prevent infinite if unable to find, but shouldn't ever happen (?)
            If Not bSearching Then
                bMatched = False
                Exit Do
            End If
        Loop

        If searchType = "t" Then
            oTextCursor.goLeft(1, bExpand)
        ElseIf searchType = "T" Then
            oTextCursor.goRight(1, bExpand)
        End If

    Else
        bMatched = False
    End If

    ' If matched, then we want to select PAST the character
    ' Else, this will counteract some weirdness. hack either way
    If Not bIsBackwards And MODE = M_VISUAL Then
        oTextCursor.goRight(1, bExpand)
    End If

    ProcessSearchKey = bMatched

End Function


Function ProcessInnerKey(oTextCursor, movementModifier, keyChar, bExpand)
    dim bMatched, searchType1, searchType2, search1, search2

    ' Setting searchType
    If movementModifier = "i" Then
        searchType1 = "T" : searchType2 = "t"
    ElseIf movementModifier = "a" Then
        searchType1 = "F" : searchType2 = "f"
    Else ' Shouldn't happen
        ProcessInnerKey = False
        Exit Function
    End If

    Select Case keyChar
	    Case "(", ")":
	        search1 = "(" : search2 = ")"
	    Case "{", "}":
	        search1 = "{" : search2 = "}"
	    Case "[", "]":
	        search1 = "[" : search2 = "}"
	    Case "<", ">":
	        search1 = "<" : search2 = ">"
	    Case "t":
	        search1 = ">" : search2 = "<"
	    Case "'":
	        search1 = "'" : search2 = "'"
	    Case """":
	        ' Matches "smart" quotes, which is default in libreoffice
	        search1 = "?" : search2 = "?"
	    Case Else:
	    	search1 = keyChar : search2 = keyChar
	End Select
	
	    dim bMatched1, bMatched2
	    bMatched1 = ProcessSearchKey(oTextCursor, searchType1, search1, False)
	    bMatched2 = ProcessSearchKey(oTextCursor, searchType2, search2, True)
	    ' Temp hack - need to search 2nd time to make da<char> work correctly
	    If (search1 = search2) and (searchType2 = "f") and bMatched1 Then
	    	bMatched2 = ProcessSearchKey(oTextCursor, searchType2, search2, True)
	    End If
	    bMatched = (bMatched1 And bMatched2)

    ProcessInnerKey = bMatched
End Function


' -----------------------
' Main Movement Function
' -----------------------
'   Default: bExpand = False, keyModifiers = 0
'   axf Need to pass oEvent to make in available to called functions (i.e. search)
Function ProcessMovementKey(keyChar, Optional bExpand, Optional keyModifiers, Optional oEvent)
    dim oTextCursor, bSetCursor, bMatched
    oTextCursor = getTextCursor()
    bMatched = True
    If IsMissing(bExpand) Then bExpand = False
    If IsMissing(keyModifiers) Then keyModifiers = 0


    ' Check for modified keys (Ctrl, Alt, not Shift)
    If keyModifiers > 1 Then
        dim bIsControl
        bIsControl = (keyModifiers = 2) or (keyModifiers = 8)

        ' Ctrl+d and Ctrl+u
        If bIsControl and keyChar = "d" Then
            getCursor().ScreenDown(bExpand)
        ElseIf bIsControl and keyChar = "u" Then
            getCursor().ScreenUp(bExpand)
        Else
            bMatched = False
        End If

        ProcessMovementKey = bMatched
        Exit Function
    End If

    ' Set global cursor to oTextCursor's new position if moved
    bSetCursor = True


    ' ------------------
    ' Movement matching
    ' ------------------

    ' ---------------------------------
    ' Special Case: Modified movements
    If getMovementModifier() <> "" Then
        Select Case getMovementModifier()
            ' f,F,t,T searching
            Case "f", "t", "F", "T":
                bMatched  = ProcessSearchKey(oTextCursor, getMovementModifier(), oEvent.keyChar, bExpand)
                LAST_SEARCH = oEvent.keyChar
            Case "i", "a":
                bMatched = ProcessInnerKey(oTextCursor, getMovementModifier(), oEvent.keyChar, bExpand)
            Case Else:
                bSetCursor = False
                bMatched = False
        End Select

        If Not bMatched Then
            bSetCursor = False
        End If
    ' ---------------------------------

    ' Search repetition
    ElseIf keyChar = "n" or keyChar = "N" Then
        If keyChar = "n" Then
            ' MsgBox("n: " & LAST_SEARCH)
            ' bMatched  = ProcessSearchKey(oTextCursor, "f", LAST_SEARCH_CHAR, bExpand)
            searchAndSet(getTextCursor(), LAST_SEARCH, False) 
        ElseIf keyChar = "N" Then
            ' MsgBox("N: " & LAST_SEARCH)
            ' bMatched  = ProcessSearchKey(oTextCursor, "F", LAST_SEARCH_CHAR, bExpand) 
			searchAndSet(getTextCursor(), LAST_SEARCH, True)
        End If
        bSetCursor = False

    ' Basic movement
    ElseIf keyChar = "l" Then
        oTextCursor.goRight(1, bExpand)

    ElseIf keyChar = "h" Then
        oTextCursor.goLeft(1, bExpand)

    ' oTextCursor.goUp and oTextCursor.goDown SHOULD work, but doesn't (I dont know why).
    ' So this is a weird hack
    ElseIf keyChar = "k" Then
        'oTextCursor.goUp(1, False)
        getCursor().goUp(1, bExpand)
        bSetCursor = False

    ElseIf keyChar = "j" Then
        'oTextCursor.goDown(1, False)
        getCursor().goDown(1, bExpand)
        bSetCursor = False
    ' ----------

    ElseIf keyChar = "^" Then
        getCursor().gotoStartOfLine(bExpand)
        bSetCursor = False
    ElseIf keyChar = "$" Then
        dim oldPos, newPos
        oldPos = getCursor().getPosition()
        getCursor().gotoEndOfLine(bExpand)
        newPos = getCursor().getPosition()

        ' If the result is at the start of the line, then it must have
        ' jumped down a line; goLeft to return to the previous line.
        '   Except for: Empty lines (check for oldPos = newPos)
        If getCursor().isAtStartOfLine() And oldPos.Y() <> newPos.Y() Then
            getCursor().goLeft(1, bExpand)
        End If

        ' maybe eventually cursorGoto... should return True/False for bsetCursor
        bSetCursor = False

    ElseIf keyChar = "w" or keyChar = "W" Then
        oTextCursor.gotoNextWord(bExpand)
    ElseIf keyChar = "b" or keyChar = "B" Then
        oTextCursor.gotoPreviousWord(bExpand)
    ElseIf keyChar = "e" Then
    	oTextCursor.goRight(1, bExpand)
        If oTextCursor.isEndOfWord() Then
            oTextCursor.gotoNextWord(bExpand)
        End If
        oTextCursor.gotoEndOfWord(bExpand)
    ElseIf keyChar = "E" Then
        oTextCursor.gotoPreviousWord(bExpand)
        oTextCursor.gotoPreviousWord(bExpand)
        oTextCursor.gotoEndOfWord(bExpand)
    ElseIf keyChar = ")" Then
        oTextCursor.gotoNextSentence(bExpand)
    ElseIf keyChar = "(" Then
        oTextCursor.gotoPreviousSentence(bExpand)
    ElseIf keyChar = "}" Then
        oTextCursor.gotoNextParagraph(bExpand)
    ElseIf keyChar = "{" Then
        oTextCursor.gotoPreviousParagraph(bExpand)

    Else
        bSetCursor = False
        bMatched = False
    End If

    ' If oTextCursor was moved, set global cursor to its position
    If bSetCursor Then
        getCursor().gotoRange(oTextCursor.getStart(), False)

        ' ---- REALLY BAD HACK
        ' I can't seem to get the View Cursor (getCursor()) to update its
        ' position without calling its own movement functions.
        ' Theoretically, the above call to gotoRange should work, but I don't
        ' know why it doesn't. Visually it works, but its X position is reset
        ' when you move lines. Bug??

		' axf Cannot just move left, check position and move right if it is different
		' because of "annotation characters". Annotations are bound to special zero-length characters
		' both ends of which are considered to have the same position. So we have to go left once,
		' check position, if it is different (in case of regular char), return right.
		' If pos is the same, this can be annotation char or the beginning of the first line, so we
		' try to move left once more. If pos is still the same, this is really the beginning of the 
		' first line and we don't need to go right at all. If not - we were trapped by annotation character
		' and need to move 2 steps right.
        dim oTempPos : oTempPos = getCursor().getPosition()
        getCursor().goLeft(1, False)        
        If Not samePos(oTempPos, getCursor().getPosition()) Then
            getCursor().goRight(1, False)
        Else
        	' Try one more step left, return 2 steps right if was trapped in annotation char
        	getCursor().goLeft(1, False)
        	If Not samePos(oTempPos, getCursor().getPosition()) Then
        		getCursor().goRight(2, False)
        	End If
        End If           
    End If

    ' If oTextCursor was moved and is in VISUAL mode, update selection
    if bSetCursor and bExpand then
        getCurrentController().Select(oTextCursor)
    end if

    ProcessMovementKey = bMatched
End Function


Sub sStartViewEventListener
	sStopViewEventListener()
	oListener = CreateUnoListener("VEListener_", "com.sun.star.document.XEventListener")
	dim oGlobalEventBroadcaster
	oGlobalEventBroadcaster = GetDefaultContext().getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
	oGlobalEventBroadcaster.addEventListener(oListener)
End Sub


Sub sStopViewEventListener
	dim oGlobalEventBroadcaster
	oGlobalEventBroadcaster = GetDefaultContext().getByName("/singletons/com.sun.star.frame.theGlobalEventBroadcaster")
	oGlobalEventBroadcaster.removeEventListener(oListener)
End Sub


Sub VEListener_notifyEvent(o)
    ' Exit if plugin is not enabled
    If MODE = M_DISABLED Then
        KeyHandler_KeyPressed = False
        Exit Sub
    End If
    
	If o is Nothing Then
		Exit Sub
	End If
	dim oSource as object
	oSource = o.Source
	If oSource is Nothing Then
		Exit Sub
	End If	
	dim bHasInterface
	bHasInterface = HasUnoInterfaces(oSource, "com.sun.star.text.XTextDocument")
	If not bHasInterface Then	
		Exit Sub
	End If
	dim oController as object
	If o.EventName = "OnFocus" Then
		reinitVibreoffice()
	ElseIf o.EventName = "OnViewCreated" Then
		oController = oSource.getCurrentController()
		If not (oController is Nothing) Then
			oController.addKeyHandler(oXKeyHandler)
		End If
	ElseIf o.EventName = "OnViewClosed" Then
		oController = oSource.getCurrentController()
		If not (oController is Nothing) Then
			oController.removeKeyHandler(oXKeyHandler)
		End If
	End If
End Sub


sub VEListener_disposing()
end sub


Sub reinitVibreoffice
  dim oTextCursor, oCurrentController
  oCurrentController = getCurrentController()
  If oCurrentController is Nothing Then
    Exit Sub
  End If

  resetMultiplier()
  setCursor()
  setTextCursor()
  gotoMode(M_NORMAL)

  ' Show terminal cursor
  oTextCursor = getTextCursor()
  If not (oTextCursor Is Nothing) Then
      cursorReset(oTextCursor)
  End If

  BUSY = False
End Sub


Sub startVibreoffice()
	If not VIBREOFFICE_STARTED Then
    	sStartXKeyHandler()
    
    	VIBREOFFICE_STARTED = True
		gotoMode(M_NORMAL)
		
		oCurrentFrame = StarDesktop().getCurrentFrame()
    End If
    
    reinitVibreoffice()
    getCurrentController().addKeyHandler(oXKeyHandler)
End Sub


Sub stopVibreoffice()
    restoreStatus()
    getCurrentController().removeKeyHandler(oXKeyHandler)
End Sub


Sub toggleVibreoffice()
    if MODE = M_DISABLED then
    	gotoMode(OLD_MODE)
    else
	    OLD_MODE = MODE
	    gotoMode(M_DISABLED)
    end if
End Sub
