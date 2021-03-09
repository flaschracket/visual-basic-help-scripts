'=========================================================================
' Written by: Samira Radan
' Date: 10/12/2020
' goal: to find out which phrases in Social Engine PHP software are not translated to Germany
'=========================================================================



inFolder = "./en"
FromLines = 1

Dim checkphrase, Englishfile
Dim enFline
Dim transFLine

'=========================================================================
 WScript.Echo "Info: Excel wird gestartet (please click OK)"
'=========================================================================
'3 data sources are needed. all of them are File
'A: our main source: german translated file
'B: checkfiles: the resources we need to check: list of english File
'C: result file: list of phrases which needed to be translated
'=========================================================================
'//
Set FSO = CreateObject("Scripting.FileSystemObject")
'//A: file of translated phrases
Set translatedsource = FSO.OpenTextFile("de.csv")

'//C: missedphrases
Set missedphrases = FSO.createTextFile("missedphrases.csv", ForWrithing, True)

'=========================================================================
'B: have list of english files in Array and loop
'=========================================================================

For Each Datei In FSO.GetFolder(inFolder).Files
 If LCase(FSO.GetExtensionName(Datei.Name)) <> "csv" Then
    WScript.Echo "Error:file extension should be csv"
      WScript.Quit
End If
Next
'-----------------------------
For Each Datei In FSO.GetFolder(inFolder).Files

Set Englishfile = FSO.OpenTextFile(Datei)

'------First: write file Name--------'
missedphrases.Write Datei.Name
'#missedphrases.Writeline
'#missedphrases.Writeline
continue = True
phraseexist = False

'------ compare file phrases with translated phrases
Do Until Englishfile.AtEndOfStream

      enFline = Englishfile.readline

      If Len(enFline) <= 0   Then
          continue = False
      Else
          enlineSplite = Split(enFline, ";")
          enphrase = enlineSplite(0)
          if InStr( enphrase , """" ) > 0 then
               enphrase = replace(enphrase, chr(34), "")
          End if
          enphrase = Trim(UCase(enphrase))

	
     Do Until translatedsource.AtEndOfStream

             transFLine = translatedsource.readline
			WScript.Echo "transfile" 
			Wscript.Echo transFLine
			translineSplite = Split(transFLine, ";")
             transphrase = translineSplite(0)
             if InStr( transphrase , """" ) > 0 then
                  transphrase = replace(transphrase, chr(34), "")
             End if
             Transphrase  = Trim(UCase(transphrase))

'===========================TEST==============
'if InStr(transphrase, "CHANGED THEIR") then
'Wscript.Echo "it has"
'WScript.Echo Transphrase & " en: " & enphrase
'End if
'============================

      If (enphrase = Transphrase) Then
                  'WScript.Echo "phrase exist"
                  phraseexist = True
      End If

      Loop
'------second: writing not translated phrase (C:missedphrases)--------'
      If (Not phraseexist AND Len(enlineSplite(0)) > 0) Then
                missedphrases.Write enlineSplite(0)
                missedphrases.Writeline
      End If


      End If

      phraseexist = False
      continue = True
      translatedsource.close
      Set translatedsource = FSO.OpenTextFile("de.csv")
Loop
Englishfile.Close
Next ' go to check new file
'=====================================================================
' Save & close file
'---------------------------------------------------------------------

missedphrases.Writeline
missedphrases.Close

'=========================================================================

WScript.Echo "Programm beendet. Vielen Dank!"

'=========================================================================

WScript.Quit
