' =========================================================================
'
' NAME: Get-WindowsProductKey.vbs
'
' COMMENT: This script retrieves the Windows product key from the registry
'          and displays it to the user in a message box.
'
' =========================================================================

Option Explicit

' Main script execution
Call GetWindowsProductKey()

Sub GetWindowsProductKey()
    ' Declare variables
    Dim wshShell, registryKeyPath, digitalProductId, productKey

    ' Create a shell object to interact with the system
    Set wshShell = CreateObject("WScript.Shell")

    ' Define the registry path where the DigitalProductId is stored
    registryKeyPath = "HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\DigitalProductId"

    ' Read the raw binary DigitalProductId from the registry
    digitalProductId = wshShell.RegRead(registryKeyPath)

    ' Decode the binary data into a readable product key
    productKey = DecodeProductKey(digitalProductId)

    ' Display the final product key to the user
    DisplayProductKey(productKey)

    ' Clean up the shell object
    Set wshShell = Nothing
End Sub

Function DecodeProductKey(digitalProductId)
    ' This function decodes the raw DigitalProductId value into a standard 
    ' 25-character product key format (XXXXX-XXXXX-XXXXX-XXXXX-XXXXX).

    ' Declare local variables for the decoding process
    Dim keyCharacters, keyOffset, i, j, currentValue, decodedKey

    ' The set of characters that a product key can contain
    keyCharacters = "BCDFGHJKMPQRTVWXY2346789"
    
    ' The byte offset within the DigitalProductId where the key data begins
    keyOffset = 52

    ' Loop backward to generate the 29 characters of the key (25 chars + 4 hyphens)
    i = 28
    Do
        currentValue = 0
        ' Loop through 15 bytes of the key data
        j = 14
        Do
            currentValue = currentValue * 256
            currentValue = digitalProductId(j + keyOffset) + currentValue
            digitalProductId(j + keyOffset) = (currentValue \ 24) And 255
            currentValue = currentValue Mod 24
            j = j - 1
        Loop While j >= 0

        i = i - 1
        decodedKey = Mid(keyCharacters, currentValue + 1, 1) & decodedKey

        ' Insert a hyphen every 5 characters
        If ((29 - i) Mod 6) = 0 And (i <> -1) Then
            i = i - 1
            decodedKey = "-" & decodedKey
        End If
    Loop While i >= 0

    DecodeProductKey = decodedKey
End Function

Sub DisplayProductKey(productKey)
    ' This subroutine displays the product key in an input box,
    ' allowing the user to select and copy the key from the text field.

    Dim messageTitle, messagePrompt

    messageTitle = "Windows 11 Current Product Key"
    messagePrompt = "Your Windows 11 product key is displayed in the text box below." & vbCrLf & vbCrLf & _
                  "Please save it in a safe place before reinstalling the operating system." & vbCrLf & vbCrLf & _
                  "You can press Ctrl+C to copy." 

    ' Use an InputBox to display the key. The text in an InputBox is selectable.
    ' The last argument sets the product key as the default text in the input field.
    InputBox messagePrompt, messageTitle, productKey
End Sub