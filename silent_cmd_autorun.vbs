'Silent Command Autorun Utility
'v1.0
'By OperaVaria, 2024.

Option Explicit

'Declare global variables:
Dim successFlag, boxTitle, oFSO, oWSS, outFile, outBeginning, outEnd, comNum

'Set successFlag, default False:
successFlag = False

'MsgBox and InputBox title string:
boxTitle = "CMD Autorun Utility"

'Set global objects:
Set oFSO = CreateObject("Scripting.FileSystemObject")
Set oWSS = CreateObject("WScript.Shell")

'Set working directory:
oWSS.CurrentDirectory = oFSO.GetParentFolderName(WScript.ScriptFullName) & "\"

'Create output file:
Set outFile = oFSO.CreateTextFile("autorun_script.vbs", True, True)

'Beginning of the template file:
outBeginning = "'Command Autorun Script" & vbCrLf & "Option Explicit" & vbCrLf &_ 
               "'Dim and Set shell object:" & vbCrLf & "Dim oShell" & vbCrLf &_
               "Set oShell = WScript.CreateObject(""WScript.shell"")" & vbCrLf
               
'End of the template file:
outEnd = "'Set object to Nothing, as tradition." & vbCrLf &_
         "Set oShell = Nothing"

Sub startConfirmMsgBox
    'Confirming process start.
    Dim startMB
    startMB = MsgBox("Do you wish to create a startup script?", 1+32+256+4096, boxTitle)
    If startMB = vbCancel Then
        exitSetup
    ElseIf startMB = vbOK Then
        askNumberInputBox
    End if
End Sub

Sub askNumberInputBox
    'MsgBoxes to set number of commands.
    Dim validFlag, numberIB, errMb
    validFlag = False
    Do While validFlag = False
        numberIB = InputBox("Enter the number of commands:", boxTitle, "Number here", 12000, 7000)        
        If numberIB = "" Then
            exitSetup
        ElseIf IsNumeric(numberIB) Then
            If numberIB <= 0 Then
                errMB = MsgBox("Error: positive numbers only!", 0+48+0+4096, boxTitle)
                validFlag = False                                    
            ElseIf numberIB <= 10 Then
                comNum = numberIB
                validFlag = True
            Else
                errMB = MsgBox("Are you sure you want to include that many commands?!", 4+32+256+4096, boxTitle)
                If errMB = vbNo Then
                    validFlag = False
                ElseIf errMB = vbYes Then
                    comNum = numberIB
                    validFlag = True
                End If
            End If
        Else            
            errMB = MsgBox("Error: input only numbers!", 0+48+0+4096, boxTitle)
            validFlag = False
        End If
    Loop  
    writeFileBeginning
End Sub

Sub writeFileBeginning
    'Write first part of the script to the output file.
    outFile.Write outBeginning
    comInpLoop
End Sub

Sub comInpLoop
    'Loop to input commands form InputBoxes. Escape double quotation character. Write command to file.
    Dim item, comIB
    For item = 1 to comNum
        comIB = InputBox("Enter command No." & item & ":", boxTitle, "%COMSPEC% /C somecommand", 12000, 7000)
        If comIB = "" Then
            'Recreate file on cancellation.
            outFile.Close
            Set outFile = oFSO.CreateTextFile("autorun_script.vbs", True, True)
            askNumberInputBox
        End If
        comIB = Replace(comIB, Chr(34), Chr(34) & Chr(34))     
        outFile.WriteLine "oShell.run " & Chr(34) & comIB & Chr(34) & ",0,false"        
    Next
    writeFileEnd
End Sub

Sub writeFileEnd
    'Write remaining part of the script to the output file.
    outFile.Write outEnd
    askAutoRun
End Sub

Sub askAutoRun
    'Success prompts. Ask to copy to autorun. Open folders.
    Dim fileSuccessMB, autorunMB, arNoMB, arYesMB, startupPath, openCurrentFolder, openStartupFolder
    successFlag = True
    fileSuccessMB = MsgBox("Script created successfully!", 0+64+0+4096, boxTitle)    
    autorunMB = MsgBox("Do you wish to copy the script to the autorun folder?", 4+32+256+4096, boxTitle)
        If autorunMB = vbNo Then
            arNoMB = MsgBox("Generated autorun file can be found in current directory." & vbCrLf &_
                            "Filename: autorun_script.vbs" , 0+64+0+4096, boxTitle)
            openCurrentFolder = oWSS.Run(oWSS.CurrentDirectory, 1, False)
        ElseIf autorunMB = vbYes Then
            startupPath = oWSS.ExpandEnvironmentStrings("%APPDATA%") &_
                          "\Microsoft\Windows\Start Menu\Programs\Startup\"
            oFSO.CopyFile "autorun_script.vbs", startupPath, True
            arYesMB = MsgBox("Script copied successfully to autorun folder." & vbCrLf &_
                             "Filename: autorun_script.vbs" , 0+64+0+4096, boxTitle)
            openStartupFolder = oWSS.Run(Chr(34) & startupPath & Chr(34), 1, False)
        End If
    exitSetup
End Sub

Sub exitSetup
    'Exiting subroutine:
    'Close outFile. Delete script if unfinished. Set objects to Nothing as tradition. Quit.
    outFile.Close
    If successFlag = False Then
        oFSO.DeleteFile("autorun_script.vbs")
    End If
    Set oFSO = Nothing
    Set oWSS = Nothing
    Set outFile = Nothing
    WScript.Quit
End Sub

'Launch Script
startConfirmMsgBox
