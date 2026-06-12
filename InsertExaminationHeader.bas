Sub InsertExaminationHeader()
'=============================================================================
' InsertExaminationHeader Macro
' Version 1.0 (Release)
' Copyright (c) 2026 Mark Gardner
'
' LICENSE:
' Permission is granted to any possessor of this macro to install and use it
' freely, provided that:
'   (1) This macro is not modified in any way;
'   (2) This copyright and license notice is preserved in full;
'   (3) Distribution to others is permitted only if the macro is provided
'       entirely free of charge and in its original, unmodified form.
'
' This software is provided "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE, AND NON-INFRINGEMENT. IN NO EVENT SHALL
' THE COPYRIGHT HOLDER BE LIABLE FOR ANY CLAIM, DAMAGES, OR OTHER LIABILITY.
'
' A new user is encouraged to read any supplied instructions in full, and to
' practice and learn the procedures before applying this to a production document.
' The brief outline of instructions given below is not complete.
'
' If there is any question whether a copy is unmodified, obtain a fresh copy with
' https://github.com/markgardnter620/insert-examination-header/releases/latest/download/InsertExaminationHeader.zip
'
' INSTRUCTIONS (abbreviated):
' See the distribution ReadMe.txt for full instructions. Briefly:
'   - Select a portion of page 1 and set a header of "temporary".
'   - Select the range of pages where Examination Headers are required;
'       precision is not required -- begin the selection anywhere on the first
'       page and end it anywhere on the final page, even if its the same page.
'   - Invoke the macro (it is convenient to assign it to a key combination
'       such as Shift+Ctrl+Alt+H).
'   - Enter the desired header text (e.g., "Smith - Direct") when prompted.
'   - The macro will set the text to all capitals or to title case as specified
'       in configuration (use the macro InsertExaminationHeaderSettings) and
'       apply headers to the selected pages, placing section breaks before and
'       after as needed. Review results carefully.
'    - Note that the macro may take several seconds to complete; the system
'       'beep' sounds twice when it is done.
'
' Installing this file per the instructions will add five (5) macros to the
' list of macros included in "Macros dialog box" (Alt+F8):
'    InsertExaminationHeader                the primary workhorse of the set
'    InsertExaminationHeaderAssignKey       assigns shortcut Shift+Ctrl+Alt+H
'    InsertExaminationHeaderClearAll        removes all section breaks & headers
'    InsertExaminationHeaderSettings        user choice center/left // title/caps
'    InsertExaminationHeaderUnassignKey     removes shortcut Shift+Ctrl_Alt_H

' Please make comments, suggestons, or disfrugalty reports to
' MarkGardnerEditor@outlook.com, subject line. "IEH" or other clear reference.
'=============================================================================

    Dim strHeaderText As String, strOriginalRaw As String, strCleanSnapshot As String
    Dim startPage As Long, endPage As Long
    Dim rStartAnchor As Range, rEndAnchor As Range
    Dim secNew As Section, secTarget As Section
    Dim currentIdx As Integer, targetIdx As Integer
    Dim firstZeroPos As Long

    Dim folderPath As String, configPath As String
    Dim useCenter As Boolean, useTitle As Boolean
    folderPath = Environ("AppData") & "\IEH_Macro\"
    configPath = folderPath & "IEH_config.txt"

    ' Check for directory presence
    If Dir(folderPath, vbDirectory) = "" Then MkDir folderPath

    ' If file is missing, build the instructor's default
    If Dir(configPath) = "" Then
        Open configPath For Output As #1
            Print #1, "CASE=TITLE"
            Print #1, "POSITION=CENTER"
        Close #1
        
        MsgBox "This is first use; a configuration file has been" & vbCrLf & _
               "constructed, defaulting to centered and title case." & vbCrLf & vbCrLf & _
               "Run 'InsertExaminationHeaderSettings' to modify these settings.", _
               vbInformation, "IEH Initialization"
        useCenter = True
        useTitle = True
    Else
        ' If file exists, parse it for settings
        Dim fNum As Integer: fNum = FreeFile ' Get a safe file handle
        Dim strLine As String

        useCenter = False
        useTitle = False
        Open configPath For Input As #fNum
        Do Until EOF(fNum)
            Line Input #fNum, strLine
            ' Search for the presence of the keywords
            If InStr(1, strLine, "CENTER", vbTextCompare) > 0 Then useCenter = True
            If InStr(1, strLine, "TITLE", vbTextCompare) > 0 Then useTitle = True
        Loop
        Close #fNum
    End If

    ' Adjacency Reconnaissance Flags
    Dim alreadyBrokenTop As Boolean: alreadyBrokenTop = False
    Dim alreadyBrokenBottom As Boolean: alreadyBrokenBottom = False

    ' Ensure something is selected
    If Selection.Start = Selection.End Then
        MsgBox "To insert Examination Headers, please first select from anywhere " & _
                "on the first page to anywhere on the last page," _
                & vbCrLf & "even if it's the same page.", vbExclamation, "Selection Required"
        Exit Sub
    End If

    ' Prevent running if the selection already contains breaks
    If Selection.Sections(1).Index <> Selection.Sections(Selection.Sections.count).Index Or _
            Selection.Sections.count > 1 Then
        MsgBox "The selected range contains one or more section breaks." _
        & vbCrLf & "You can select within an existing range, or in as-yet" _
        & " unsectioned portions of the document." _
        , vbCritical, "Existing Section Breaks Detected"
        Exit Sub
    End If

' ==========================================================================
    ' GET EXISTING HEADER (Forensic Architecture Parser v0.8)
    ' ==========================================================================
    'Dim strOriginalRaw As String
    'Dim strCleanSnapshot As String
    Dim strVolumeText As String
    Dim strOutput As String
    Dim strMarks As String
    
    Dim onePos As Long
    Dim leftModeTabPos As Long
    'Dim firstZeroPos As Long
    Dim secondZeroPos As Long
    Dim searchStart As Long
    Dim adjacentChar As String
    Dim wasCentered As Boolean
    
    ' Capture and clean trailing carriage returns/whitespace defensively
    strOriginalRaw = Selection.Sections(1).Headers(wdHeaderFooterPrimary).Range.text
    strOriginalRaw = Replace(strOriginalRaw, vbTab, "0") ' need introduced in Word update, sigh.
    strOriginalRaw = RTrim(strOriginalRaw)
    If Right(strOriginalRaw, 1) = vbCr Then
        strOriginalRaw = Left(strOriginalRaw, Len(strOriginalRaw) - 1)
    End If
    
    ' Establish the absolute page field anchor token (0x31) working backwards
    onePos = InStrRev(strOriginalRaw, "1")
    
    ' Fallback if the structural page token is completely missing
    If onePos = 0 Then
        strCleanSnapshot = Trim(strOriginalRaw) & "?"
        GoTo AssemblyComplete
    End If
    
' ==========================================================================
    ' Evaluate layout structural scenarios
    ' ==========================================================================
    ' First, execute the forensic backward-scan to find the layout tab boundary
    leftModeTabPos = 0
    If Len(strOriginalRaw) > 1 Then
        searchStart = onePos - 1
        leftModeTabPos = InStrRev(strOriginalRaw, "0", searchStart)
    End If
    
' Hop backward over numeric digits or hyphens (e.g., "10-10") to clear data zeros
    Do While leftModeTabPos > 1
        adjacentChar = Mid$(strOriginalRaw, leftModeTabPos - 1, 1)
        
        ' CRITICAL GUARD: Only treat as data if it's not the structural zero at position 1
        If (IsNumeric(adjacentChar) Or adjacentChar = "-") And (leftModeTabPos - 1 > 1) Then
            leftModeTabPos = InStrRev(strOriginalRaw, "0", leftModeTabPos - 1)
        Else
            Exit Do
        End If
    Loop
    
    ' Determine geometry by testing the position of the true layout tab
    firstZeroPos = InStr(1, strOriginalRaw, "0")
    
    If firstZeroPos = 1 And leftModeTabPos > 1 Then
        ' === CASE: TRULY CENTERED TARGET (Has leading tab AND an internal layout tab) ===
        wasCentered = True
        strMarks = "???"
        
        ' Look for an immediate second alignment tab token for naked templates
        secondZeroPos = InStr(2, strOriginalRaw, "0")
        
        If secondZeroPos = 2 Then
            ' Scenario: "00[vol]1" -> Empty text centered header
            strCleanSnapshot = ""
            
            ' Extract volume text directly using our clean index math optimization
            If onePos - secondZeroPos > 1 Then
                strVolumeText = Mid$(strOriginalRaw, 3, onePos - secondZeroPos - 1)
            Else
                strVolumeText = ""
            End If
        Else
            ' Scenario: "0[text]0[vol]1" -> Centered text exists
            ' Extract text between the leading tab (pos 1) and the trailing layout tab
            If leftModeTabPos > 2 Then
                strCleanSnapshot = Mid$(strOriginalRaw, 2, leftModeTabPos - 2)
            Else
                strCleanSnapshot = ""
            End If
            
            ' Isolate volume text between the internal layout tab and the page token
            If onePos - leftModeTabPos > 1 Then
                strVolumeText = Mid$(strOriginalRaw, leftModeTabPos + 1, onePos - leftModeTabPos - 1)
            Else
                strVolumeText = ""
            End If
        End If
    Else
        ' CASE: Left-Justified/Manual Target
        wasCentered = False
        strMarks = "??"
       
        ' Catch literal "01" empty left-justified setup
        If Left(strOriginalRaw, 2) = "01" Then
            strCleanSnapshot = ""
            strVolumeText = ""
        Else
            ' Find the layout tab before the page number by searching backwards
            leftModeTabPos = 0
            If Len(strOriginalRaw) > 1 Then
                searchStart = onePos - 1
                leftModeTabPos = InStrRev(strOriginalRaw, "0", searchStart)
            End If
            ' Hop backward over hyphenated/numeric data characters
            Do While leftModeTabPos > 1
                adjacentChar = Mid(strOriginalRaw, leftModeTabPos - 1, 1)
                If IsNumeric(adjacentChar) Or adjacentChar = "-" Then
                    leftModeTabPos = InStrRev(strOriginalRaw, "0", leftModeTabPos - 1)
                Else
                    Exit Do
                End If
            Loop
            
            If leftModeTabPos > 0 Then
                ' Structural alignment tab confirmed
                strMarks = "??"
                strCleanSnapshot = Left(strOriginalRaw, leftModeTabPos - 1)
                
                If onePos - leftModeTabPos > 1 Then
                    strVolumeText = Mid(strOriginalRaw, leftModeTabPos + 1, onePos - leftModeTabPos - 1)
                Else
                    strVolumeText = ""
                End If
            Else
                ' === NO ALIGNMENT TABS FOUND (Manual justification / Tab stops) ===
                If strMarks <> "??" Then
                    strMarks = "?"
                End If
                
                If onePos = 1 Then
                    ' Scenario: "1" -> Page number only, right-justified paragraph
                    strCleanSnapshot = ""
                    strVolumeText = ""
                Else
                    ' Check for manual tab stops (Chr(9))
                    If Mid(strOriginalRaw, onePos - 1, 1) = Chr(9) Then
                        ' Scenario: "[tab]1"
                        strCleanSnapshot = Left(strOriginalRaw, onePos - 1)
                        strVolumeText = ""
                    ElseIf InStr(1, strOriginalRaw, Chr(9)) > 0 Then
                        ' Scenario: "[tab][vol]1"
                        Dim tabLoc As Long
                        tabLoc = InStr(1, strOriginalRaw, Chr(9))
                        strCleanSnapshot = Left(strOriginalRaw, tabLoc)
                        strVolumeText = Mid(strOriginalRaw, tabLoc + 1, onePos - tabLoc - 1)
                    Else
                        ' Scenario: "[vol]1" -> Pure data prefix preceding the page anchor
                        strCleanSnapshot = ""
                        strVolumeText = Left(strOriginalRaw, onePos - 1)
                    End If
                End If
            End If
        End If
    End If

AssemblyComplete:
    ' Final normalization and string packaging
    strCleanSnapshot = Trim(strCleanSnapshot)
    strVolumeText = Trim(strVolumeText)
    
    strOutput = strCleanSnapshot & strMarks & strVolumeText
    ' ==========================================================================

    ' Combined User Input
    Do
        strHeaderText = InputBox("InsertExaminationHeader v1p0." _
                     & vbCrLf & vbCrLf & "TO PROCEED, enter the header text below " _
                     & "(e.g., ""Smith - Direct""; it will be set correctly)." & vbCrLf & vbCrLf _
                     & "TO CANCEL, enter 'No' or press Cancel.", "Enter Header Text")
        
        If StrPtr(strHeaderText) = 0 Or UCase(strHeaderText) = "NO" Then Exit Sub
        If Trim(strHeaderText) <> "" Then Exit Do
    Loop While True

    strHeaderText = UCase(strHeaderText)
    
    On Error GoTo ErrorHandler

	' Initialize the Undo Record
    Dim objUndo As UndoRecord
    Set objUndo = Application.UndoRecord

	Application.ScreenUpdating = False
    DoEvents: ActiveDocument.Repaginate: DoEvents
    objUndo.StartCustomRecord ("Insert Examination Header")

    ' IDENTIFY BOUNDARIES
    Dim tempRange As Range
    Set tempRange = Selection.Range
    tempRange.Collapse wdCollapseStart
    startPage = tempRange.Information(wdActiveEndAdjustedPageNumber)
    tempRange.End = Selection.End
    tempRange.Collapse wdCollapseEnd
    endPage = tempRange.Information(wdActiveEndAdjustedPageNumber)

    ' RECONNAISSANCE
    If startPage > 1 Then
        Set rStartAnchor = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=startPage)
        If rStartAnchor.Sections(1).Index <> rStartAnchor.Previous(unit:=wdCharacter, count:=1).Sections(1).Index Then
            alreadyBrokenTop = True
        End If
    End If

    Set rEndAnchor = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=endPage + 1)
    If rEndAnchor.Sections(1).Index <> rEndAnchor.Previous(unit:=wdCharacter, count:=1).Sections(1).Index Then
        alreadyBrokenBottom = True
    End If

    ' --- SECTION SURGERY ---
    Dim breakRetry As Integer

    ' Handle the "Top" break
    If startPage > 1 And Not alreadyBrokenTop Then
        Set rStartAnchor = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=startPage)
        rStartAnchor.MoveStart unit:=wdCharacter, count:=-1
        rStartAnchor.Collapse wdCollapseStart
        
        If rStartAnchor.Characters(1).text = " " Or rStartAnchor.Characters(1).text = vbCr Then
            rStartAnchor.MoveStart unit:=wdCharacter, count:=1
            rStartAnchor.Collapse wdCollapseStart
        End If
     
        If rStartAnchor.Sections(1).Index = rStartAnchor.Next(unit:=wdCharacter, count:=1).Sections(1).Index Then
            currentIdx = rStartAnchor.Sections(1).Index
            Dim isMidParaTop As Boolean: isMidParaTop = (rStartAnchor.Previous(unit:=wdCharacter, count:=1).text <> vbCr)
            
            rStartAnchor.InsertBreak Type:=wdSectionBreakNextPage
            
            ' Spin Lock for immediate isolation
            For breakRetry = 1 To 100
                ActiveDocument.Sections(currentIdx + 1).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
                DoEvents
                If ActiveDocument.Sections(currentIdx + 1).Headers(wdHeaderFooterPrimary).LinkToPrevious = False Then Exit For
            Next breakRetry
            
            ActiveDocument.Repaginate: DoEvents
            
            If isMidParaTop Then ActiveDocument.Sections(currentIdx + 1).Range.Paragraphs(1).FirstLineIndent = 0
        End If
    End If

    ' Handle the "Bottom" break
    If Not alreadyBrokenBottom Then
        Set rEndAnchor = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=endPage + 1)
        rEndAnchor.MoveStart unit:=wdCharacter, count:=-1
        rEndAnchor.Collapse wdCollapseStart
        
        If rEndAnchor.Characters(1).text = " " Or rEndAnchor.Characters(1).text = vbCr Then
            rEndAnchor.MoveStart unit:=wdCharacter, count:=1
            rEndAnchor.Collapse wdCollapseStart
        End If
     
        If rEndAnchor.Sections(1).Index = rEndAnchor.Next(unit:=wdCharacter, count:=1).Sections(1).Index Then
            currentIdx = rEndAnchor.Sections(1).Index
            Dim isMidParaBot As Boolean: isMidParaBot = (rEndAnchor.Previous(unit:=wdCharacter, count:=1).text <> vbCr)
            
            rEndAnchor.InsertBreak Type:=wdSectionBreakNextPage
            
            ' Spin Lock for immediate isolation
            For breakRetry = 1 To 100
                ActiveDocument.Sections(currentIdx + 1).Headers(wdHeaderFooterPrimary).LinkToPrevious = False
                DoEvents
                If ActiveDocument.Sections(currentIdx + 1).Headers(wdHeaderFooterPrimary).LinkToPrevious = False Then Exit For
            Next breakRetry
            
            ActiveDocument.Repaginate: DoEvents

            If isMidParaBot Then ActiveDocument.Sections(currentIdx + 1).Range.Paragraphs(1).FirstLineIndent = 0
        End If
    End If
    
    DoEvents: ActiveDocument.Repaginate: DoEvents

    ' --- FINAL HEADER APPLICATION ---
    Set rStartAnchor = ActiveDocument.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, count:=startPage)
    Set secTarget = rStartAnchor.Sections(1)
    targetIdx = secTarget.Index

    ' Apply new header using hardened builder
    Dim isStandardHeader As Boolean
    isStandardHeader = Build_IEH_Header(secTarget, strHeaderText, strVolumeText, useCenter, useTitle)

    If isStandardHeader Then
        ' Cleanup Preceding Section Boundary
        If Not alreadyBrokenTop And targetIdx > 1 Then
            ' Pass: Target Section, Header Layer, and the Forensic Signature String
            Restore_IEH_Header ActiveDocument.Sections(targetIdx - 1), wdHeaderFooterPrimary, strOutput
        End If
    
        ' Cleanup Succeeding Section Boundary
        If Not alreadyBrokenBottom And targetIdx < ActiveDocument.Sections.count Then
            ' Pass: Target Section, Header Layer, and the Forensic Signature String
            Restore_IEH_Header ActiveDocument.Sections(targetIdx + 1), wdHeaderFooterPrimary, strOutput
        End If
    End If

    ' Close the record normally
    objUndo.EndCustomRecord
    Application.ScreenUpdating = True
    DoEvents: ActiveDocument.Repaginate: DoEvents
    Call FinalSignal
    Exit Sub

ErrorHandler:
    ' Ensure the environment is restored regardless of the error
    On Error Resume Next ' Prevent nested errors in the handler
    objUndo.EndCustomRecord
    Application.ScreenUpdating = True

    ' Handle specific known issues
    Select Case Err.Number
        Case 0 ' No error, just exit
        Case 4198 ' Command failed
            MsgBox "The operation was interrupted or could not be completed." & vbCrLf _
                & "Check for header damage. Is file read-only?", vbExclamation
        Case Else ' Unexpected errors
            MsgBox "An unexpected error occurred (Error " & Err.Number & ")." & vbCrLf & _
                   "Description: " & Err.Description, vbCritical, "InsertExaminationHeader Error"
    End Select
    ' Clear the error
    Err.Clear
End Sub

' Build and insert the specified header with Literal-Zero Sniffing and Prefix Preservation
Function Build_IEH_Header(targetSection As Section, txt As String, volPrefix As String, bCenter As Boolean, bTitle As Boolean) As Boolean
    Dim hRng As Range, rPage As Range, rStart As Range, rTilde As Range
    Dim hType As WdHeaderFooterIndex: hType = wdHeaderFooterPrimary
    Dim breakRetry As Integer
    Dim testTxt As String
    Dim firstZeroPos As Long, lastZeroPos As Long, onePos As Long

    Build_IEH_Header = True

    If bTitle Then txt = StrConv(txt, vbProperCase)

    ' TABLE SAFETY CHECK
    ' Protects agency templates from being deleted.
    If targetSection.Headers(hType).Range.Tables.count > 0 Then
        MsgBox "Table-based header detected. Section isolation is complete, " & _
               "but table entries have not been updated; please edit manually.", _
               vbInformation, "Agency Template Protection"
                Build_IEH_Header = False
        Exit Function
    End If

    ' ISOLATION
    ' Loop ensures LinkToPrevious is truly False before proceeding.
    For breakRetry = 1 To 50
        targetSection.Headers(hType).LinkToPrevious = False
        DoEvents
        If targetSection.Headers(hType).LinkToPrevious = False Then Exit For
    Next breakRetry

    ' BUILDING (The Construction Engine)
    Set hRng = targetSection.Headers(hType).Range
    hRng.MoveEnd wdCharacter, -1
    hRng.InsertBefore "#"
    hRng.Start = hRng.Start + 1
    hRng.Delete

    ' Note: 'txt' here should be your processed finalTxt (Title or Caps)
    'hRng.text = txt & "~" & volPrefix

    Set hRng = targetSection.Headers(hType).Range
    hRng.MoveEnd wdCharacter, -1 ' Exclude paragraph mark
    hRng.InsertBefore (txt & "~" & volPrefix)

    ' Find "~#" and replace it with "~" inside the header range
    With hRng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .text = "#"
        .Replacement.text = ""
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
        .Execute Replace:=wdReplaceOne
    End With
    
    ' Add the Page Field at the end
    Set rPage = hRng.Duplicate
    rPage.Collapse wdCollapseEnd
    targetSection.Headers(hType).Range.Fields.Add Range:=rPage, _
        Type:=wdFieldPage, PreserveFormatting:=False
    
    ' --- THE DYNAMIC POSITION GATE ---
    ' Only add the Center Alignment Tab if bCenter is TRUE
    If bCenter Then
        Set hRng = targetSection.Headers(hType).Range
        Set rStart = hRng.Duplicate
        rStart.Collapse wdCollapseStart
        rStart.InsertAlignmentTab Alignment:=wdCenter, RelativeTo:=wdMargin
    End If
    
    ' Swap the Tilde for the Right Alignment Tab (Always needed for the Page Number)
    Set rTilde = targetSection.Headers(hType).Range
    With rTilde.Find
        .text = "~"
        .Forward = True
        .Wrap = wdFindStop
        If .Execute Then
            rTilde.InsertAlignmentTab Alignment:=wdRight, RelativeTo:=wdMargin
        End If
    End With
End Function

Public Sub Restore_IEH_Header(targetSection As Section, hType As WdHeaderFooterIndex, forensicString As String)
    Dim hRng As Range
    Dim cleanText As String
    Dim volText As String
    Dim templateTxt As String
    Dim marksPos As Long
    Dim marksCount As Integer
    Dim i As Integer
    Dim localForensicString As String
    
    ' Locate the Question Mark Splitter
    localForensicString = forensicString
    marksPos = InStr(localForensicString, "?")
    If marksPos = 0 Then Exit Sub ' Safety escape if string is malformed
    
    ' Count the marks to determine geometry
    marksCount = 0
    For i = marksPos To Len(localForensicString)
        If Mid(localForensicString, i, 1) = "?" Then
            marksCount = marksCount + 1
        Else
            Exit For
        End If
    Next i

    ' Extract the clean data payload segments
    cleanText = Trim(Left(localForensicString, marksPos - 1))
    volText = Trim(Mid(localForensicString, marksPos + marksCount))
    
    If Len(volText) > 0 Then
        ' Slice off everything after the question marks
        localForensicString = Left(localForensicString, marksPos + marksCount - 1)
    End If
    
    ' Break link to previous section (Skip if it's Section 1)
    If targetSection.Index > 1 Then
        If targetSection.Headers(hType).LinkToPrevious = True Then
            targetSection.Headers(hType).LinkToPrevious = False
            DoEvents ' Ensure Word updates internal document DOM
        End If
    End If
    
    ' Clear the Bench completely (Preserving the Paragraph Container)
    Set hRng = targetSection.Headers(hType).Range
    hRng.MoveEnd wdCharacter, -1
    hRng.InsertBefore "#"
    hRng.Start = hRng.Start + 1
    hRng.Delete
    
    ' Reset character properties but leave the paragraph format container alone
    Set hRng = targetSection.Headers(hType).Range
    hRng.Font.Reset
    
    ' Assemble the Raw Text Template String using Tildes
    Select Case marksCount
        Case 3
            ' === WAS CENTERED MODE ===
            If Len(volText) > 0 Then
                templateTxt = "~" & cleanText & "~" & volText
            Else
                templateTxt = "~" & cleanText & "~"
            End If
            
        Case 2
            ' === WAS LEFT-JUSTIFIED MODE ===
            If Len(volText) > 0 Then
                templateTxt = cleanText & "~" & volText
            Else
                templateTxt = cleanText & "~"
            End If
            
        Case 1
            ' === WAS MANUAL TAB STOP / FORCED RIGHT JUSTIFICATION ===
            templateTxt = cleanText & volText
            
            ' Apply alignment formatting if it was purely forced right
            If InStr(localForensicString, Chr(9)) = 0 And InStr(localForensicString, "0") = 0 Then
                targetSection.Headers(hType).Range.ParagraphFormat.Alignment = wdAlignParagraphRight
            End If
    End Select
    
    ' Print the Flat Template Text into the Header
    Set hRng = targetSection.Headers(hType).Range
    hRng.MoveEnd wdCharacter, -1
    hRng.text = templateTxt
    
    ' Append the Live Page Field Token BEFORE swapping placeholders
    Set hRng = targetSection.Headers(hType).Range
    hRng.MoveEnd wdCharacter, -1 ' Position right before the trailing ¶
    hRng.Collapse wdCollapseEnd ' Drop anchor right at the end of text payload
    
    targetSection.Headers(hType).Range.Fields.Add Range:=hRng, Type:=wdFieldPage
    
    ' Process Placeholders from Left to Right
    ' --- First Tilde Pass (Centered Mode's Leading Tab) ---
    If marksCount = 3 Then
        Set hRng = targetSection.Headers(hType).Range
        If InStr(hRng.text, "~") > 0 Then
            hRng.Find.Execute FindText:="~"
            hRng.text = ""
            hRng.InsertAlignmentTab Alignment:=wdCenter, RelativeTo:=wdMargin
        End If
    End If
    
    ' --- Second Tilde Pass (The Main Trailing/Right Tab) ---
    If marksCount >= 2 Then
        Set hRng = targetSection.Headers(hType).Range
        If InStr(hRng.text, "~") > 0 Then
            hRng.Find.Execute FindText:="~"
            hRng.text = ""
            hRng.InsertAlignmentTab Alignment:=wdRight, RelativeTo:=wdMargin
        End If
    End If
End Sub
    
Sub FinalSignal()
    Dim finishTime As Double
    
    Beep
    finishTime = Timer + 0.5

    Do While Timer < finishTime
        DoEvents
    Loop
    
    Beep
End Sub

Public Sub InsertExaminationHeaderClearAll()
    Dim doc As Document
    Set doc = ActiveDocument
    
    ' Upfront Safety Verification Dialog
    Dim userChoice As VbMsgBoxResult
    userChoice = MsgBox("This procedure will remove all headers from this document, and remove all section breaks. Do you want to proceed?", _
                        vbYesNo + vbQuestion + vbDefaultButton2, _
                        "Verify IEH Total Clear")
    
    If userChoice <> vbYes Then Exit Sub
    
    ' Fast execution setup
    Application.ScreenUpdating = False
    
    Dim i As Long
    Dim breakRange As Range
    Dim safetyCheckRange As Range
    Dim boundaryChar As String
    
    ' PHASE 1: Collapse the structural walls first (Bottom-to-Top)
    For i = doc.Sections.count To 2 Step -1
        Set breakRange = doc.Sections(i).Range
        breakRange.Collapse wdCollapseStart
        breakRange.MoveStart wdCharacter, -1
        
        If breakRange.Characters(1).text = Chr(12) Or doc.Sections(i).Index > 1 Then
            
            ' Character isolation check
            Set safetyCheckRange = breakRange.Duplicate
            safetyCheckRange.Collapse wdCollapseStart
            safetyCheckRange.MoveStart wdCharacter, -1
            boundaryChar = safetyCheckRange.text
            
            ' Drop a safety space if text blocks are directly touching the break
            If boundaryChar <> Chr(32) And boundaryChar <> Chr(13) Then
                safetyCheckRange.Collapse wdCollapseEnd
                safetyCheckRange.text = " "
                
                ' CRITICAL FIX: Re-sync breakRange to ensure it hasn't been
                ' pushed forward by the newly inserted space.
                Set breakRange = doc.Sections(i).Range
                breakRange.Collapse wdCollapseStart
                breakRange.MoveStart wdCharacter, -1
            End If
            
            ' Drop the wall cleanly
            breakRange.Delete
        End If
    Next i
    
    ' PHASE 2: Clean the slate exactly ONCE for the whole document
    ' Now that there is only Section 1 left, this single call resets everything.
    Call Restore_IEH_Header(doc.Sections(1), wdHeaderFooterPrimary, "???")
    
    ' Wrap up and notify
    Application.ScreenUpdating = True
    
    Call FinalSignal
    
    MsgBox "All examination headers removed.", _
           vbInformation, "IEH Clear All Complete"
End Sub

Sub InsertExaminationHeaderSettings()
    Dim folderPath As String, configPath As String
    Dim finalPos As String, finalCase As String
    Dim response As VbMsgBoxResult
    
    folderPath = Environ("AppData") & "\IEH_Macro\"
    configPath = folderPath & "IEH_config.txt"
    
    ' Step One: Determine Position
    response = MsgBox("Choose Header Position:" & vbCrLf & vbCrLf & _
                      "Click YES for CENTERED" & vbCrLf & _
                      "Click NO for LEFT-ALIGNED", _
                      vbYesNoCancel + vbQuestion, "IEH Layout Configuration")
                      
    If response = vbCancel Then Exit Sub
    If response = vbYes Then finalPos = "CENTER" Else finalPos = "LEFT"
    
    ' Step Two: Determine Casing
    response = MsgBox("Choose Text Case Configuration:" & vbCrLf & vbCrLf & _
                      "Click YES for TITLE CASE (Mixed)" & vbCrLf & _
                      "Click NO for UPPER CASE (ALL CAPS)", _
                      vbYesNoCancel + vbQuestion, "IEH Style Configuration")
                      
    If response = vbCancel Then Exit Sub
    If response = vbYes Then finalCase = "TITLE" Else finalCase = "UPPER"
    
    ' Commit to file
    On Error GoTo FileError
    Open configPath For Output As #1
        Print #1, "POSITION=" & finalPos
        Print #1, "CASE=" & finalCase
    Close #1
    
    MsgBox "Configuration Updated Successfully:" & vbCrLf & vbCrLf & _
           "Position: " & finalPos & vbCrLf & _
           "Casing: " & finalCase, vbInformation, "IEH Settings Saved"
    Exit Sub

FileError:
    MsgBox "Error writing to configuration file. Ensure the folder exists.", vbCritical, "File Error"
End Sub

Sub InsertExaminationHeaderAssignKey()
    ' Assigns the Shift+Ctrl+Alt+H shortcut to the primary IEH macro globally
    
    On Error GoTo ErrorHandler
    
    ' Ensure the shortcut is saved to the global Normal template
    CustomizationContext = NormalTemplate
    
    ' Add the key binding
    KeyBindings.Add _
        KeyCategory:=wdKeyCategoryMacro, _
        command:="InsertExaminationHeader", _
        KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyAlt, wdKeyH)
        
    MsgBox "Shortcut key 'Shift+Ctrl+Alt+H' successfully assigned to InsertExaminationHeader.", _
           vbInformation, "IEH Shortcut Assignment"
           
    Exit Sub

ErrorHandler:
    MsgBox "Unable to assign shortcut key automatically." & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "Assignment Error"
End Sub

Sub InsertExaminationHeaderUnassignKey()
    ' Forcefully unbinds anything currently attached to that key combination
    ' across BOTH the document and the global template.
    
    ' Clear it from the active document context
    CustomizationContext = ActiveDocument
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyAlt, wdKeyH)).Clear
    
    ' Clear it from the global Normal template context
    CustomizationContext = NormalTemplate
    FindKey(BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyAlt, wdKeyH)).Clear
    
    MsgBox "InsertExaminationHeader is no longer connected to key combination.", vbInformation
End Sub
