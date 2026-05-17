Option Explicit

Dim fso, shell, inputFile, outputFile, mode, inputFilename, outputFilename, inputText, outputText, key
Dim folderObj, chosenFolder, fullPath, customKey, extChoice, outputFolderObj, outputFolderPath, finalOutName, customName
Dim keyConfirm, openConfirm

Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("Shell.Application")

' 1. Ask the user for the input filename
inputFilename = InputBox("Enter the exact name of the text file to read (e.g., secret.txt):", "File Selection")
inputFilename = Trim(inputFilename)

If inputFilename = "" Then WScript.Quit

' 2. Loop if input file is not found
fullPath = inputFilename
Do While Not fso.FileExists(fullPath)
    MsgBox "Error: '" & inputFilename & "' not found in the current folder! Please select the folder where it is located.", 48, "File Not Found"
    Set folderObj = shell.BrowseForFolder(0, "Select the folder containing " & inputFilename, 0)
    If folderObj Is Nothing Then WScript.Quit
    chosenFolder = folderObj.Self.Path
    If Right(chosenFolder, 1) <> "\" Then chosenFolder = chosenFolder & "\"
    fullPath = chosenFolder & inputFilename
Loop

' 3. Ask for a custom encryption key with confirmation loop
Do
    customKey = InputBox("Enter your secret encryption key:" & vbCrLf & "(Leave blank to use default 'Elytrasaurus')", "Cipher Key Setup")
    customKey = Trim(customKey)
    
    If customKey = "" Then
        key = "Elytrasaurus"
        Exit Do
    Else
        key = customKey
        ' Verification step to prevent typos
        keyConfirm = MsgBox("Your key is set to: " & key & vbCrLf & "Is this correct?", 4 + 32, "Confirm Key")
        If keyConfirm = 6 Then Exit Do ' 6 means "Yes"
    End If
Loop

' 4. Ask the user if they want to Encrypt or Decrypt
mode = InputBox("Type E to ENCRYPT or D to DECRYPT:", "Vigenere Tool", "E")
mode = UCase(Trim(mode))
If mode <> "E" And mode <> "D" Then
    MsgBox "Invalid choice! Please type E or D.", 16, "Error"
    WScript.Quit
End If

' 5. Pick where to save the new file
MsgBox "Please select the destination folder where you want to save the new file.", 64, "Select Output Directory"
Set outputFolderObj = shell.BrowseForFolder(0, "Select Destination Folder", 0)
If outputFolderObj Is Nothing Then WScript.Quit
outputFolderPath = outputFolderObj.Self.Path
If Right(outputFolderPath, 1) <> "\" Then outputFolderPath = outputFolderPath & "\"

' 6. Ask the user exactly what to name the final file
customName = InputBox("Enter a name for your new file (Do not include the extension like .txt):", "Name Your File", "output_data")
customName = Trim(customName)
If customName = "" Then customName = "output_data"

' 7. Extension Choice based on Encryption/Decryption mode
If mode = "E" Then
    extChoice = InputBox("Choose an output file extension:" & vbCrLf & "1 = Normal (.txt)" & vbCrLf & "2 = Disguised/Hidden (.dat)", "File Extension Choice", "1")
    If extChoice = "2" Then
        finalOutName = customName & ".dat"
    Else
        finalOutName = customName & ".txt"
    End If
Else
    finalOutName = customName & ".txt"
End If

' Combine path and your custom name
outputFilename = outputFolderPath & finalOutName

' 8. Read the input file (with an immediate emptiness check)
Set inputFile = fso.OpenTextFile(fullPath, 1) ' 1 = ForReading

' If it's already at the end of the stream right when opening, the file is empty
If inputFile.AtEndOfStream Then
    inputFile.Close
    MsgBox "Error: Document Empty!" & vbCrLf & "Please add some text to your file and try again.", 16, "Empty File"
    WScript.Quit
End If

' If we made it past the check, it's safe to read!
inputText = inputFile.ReadAll
inputFile.Close

' 9. Process the text based on the chosen mode
If mode = "E" Then
    outputText = VigenereEncrypt(inputText, key)
    MsgBox "Success! Encrypted version created at: " & vbCrLf & outputFilename, 64, "Done"
Else
    outputText = VigenereDecrypt(inputText, key)
    MsgBox "Success! Decrypted version created at: " & vbCrLf & outputFilename, 64, "Done"
End If

' 10. Create and write to the new file at the custom destination
Set outputFile = fso.CreateTextFile(outputFilename, True) ' True = Overwrite
outputFile.Write outputText
outputFile.Close

' 11. Optional Auto-Open Feature
openConfirm = MsgBox("Would you like to open the new file in Notepad right now?", 4 + 32, "Open File")
If openConfirm = 6 Then ' 6 means "Yes"
    CreateObject("WScript.Shell").Run "notepad.exe """ & outputFilename & """"
End If


' --- Vigenere Encryption Function ---
Function VigenereEncrypt(text, cipherKey)
    Dim i, charCode, keyIndex, keyChar, keyShift, outText, textChar
    cipherKey = LCase(cipherKey)
    keyIndex = 0
    outText = ""
    
    For i = 1 To Len(text)
        textChar = Mid(text, i, 1)
        charCode = Asc(textChar)
        
        If charCode >= 65 And charCode <= 90 Then
            keyChar = Mid(cipherKey, (keyIndex Mod Len(cipherKey)) + 1, 1)
            keyShift = Asc(keyChar) - 97
            textChar = Chr(((charCode - 65 + keyShift) Mod 26) + 65)
            keyIndex = keyIndex + 1
        ElseIf charCode >= 97 And charCode <= 122 Then
            keyChar = Mid(cipherKey, (keyIndex Mod Len(cipherKey)) + 1, 1)
            keyShift = Asc(keyChar) - 97
            textChar = Chr(((charCode - 97 + keyShift) Mod 26) + 97)
            keyIndex = keyIndex + 1
        End If
        outText = outText & textChar
    Next
    VigenereEncrypt = outText
End Function

' --- Vigenere Decryption Function ---
Function VigenereDecrypt(text, cipherKey)
    Dim i, charCode, keyIndex, keyChar, keyShift, outText, textChar
    cipherKey = LCase(cipherKey)
    keyIndex = 0
    outText = ""
    
    For i = 1 To Len(text)
        textChar = Mid(text, i, 1)
        charCode = Asc(textChar)
        
        If charCode >= 65 And charCode <= 90 Then
            keyChar = Mid(cipherKey, (keyIndex Mod Len(cipherKey)) + 1, 1)
            keyShift = Asc(keyChar) - 97
            textChar = Chr(((charCode - 65 - keyShift + 26) Mod 26) + 65)
            keyIndex = keyIndex + 1
        ElseIf charCode >= 97 And charCode <= 122 Then
            keyChar = Mid(cipherKey, (keyIndex Mod Len(cipherKey)) + 1, 1)
            keyShift = Asc(keyChar) - 97
            textChar = Chr(((charCode - 97 - keyShift + 26) Mod 26) + 97)
            keyIndex = keyIndex + 1
        End If
        outText = outText & textChar
    Next
    VigenereDecrypt = outText
End Function