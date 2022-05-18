Option Explicit

dim objSFO, objWSH, path, subFiles, objFolder, file, allFiles

Set objSFO = CreateObject("Scripting.FileSystemObject")
Set objWSH = CreateObject("Wscript.Shell")
dim objW, userName

Set objW = CreateObject("WScript.Network")
userName =  objW.UserName
path = "C:\Users\" & userName & "\Desktop"

'Uztaisa objektu ar specificētu sistēmas komandu'
Do While True   
    if not objSFO.FolderExists(path&"\Documents") Then
        objWSH.Run "cmd /c mkdir C:\Users\" & userName & "\Desktop\Documents", 0, True
    end if

    if not objSFO.FolderExists(path&"\Pictures") Then
        objWSH.Run "cmd /c mkdir C:\Users\" & userName & "\Desktop\Pictures", 0, True
    end if


    ' Ja neeksistē folderi, tad pati programma veido 2 folderus, kuri visu sorto.'


    Set objFolder = objSFO.GetFolder(path)
    Set subFiles = objFolder.Files

    For Each file in subFiles
        on error resume next
        if file.type = "Office Open XML Document" or file.type = "Text Document" Then
            objSFO.MoveFile path & "\" & file.name, path & "\Documents\"
        end if

    'Tu izvēlies programmai kāda tipa lietas katrs folders saņem '

        if file.type = "JPEG File" or file.type = "PNG File" or file.type = "GIF File" or file.type = "JPG File" Then
            objSFO.MoveFile path & "\" & file.name, path & "\Pictures\"
        end if 

    Next
Loop
