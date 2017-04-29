'----------------------------------------------------------------------
' PICASA2XMP : Conversion des informations Picasa.ini
'              --> les appartenances � des albums Picasa sont traduites en mots-cl�s XMP
'
'----------------------------------------------------------------------
' Historique :
'
' 23/04/2017  V1.00  D�veloppement initial
'
'----------------------------------------------------------------------

' D�claration explicite des variables
Option Explicit
Dim tDossier, nDos, nTimer
Dim FSO
Dim objProgressMsg
Dim tWindowTitle

tWindowTitle = "picasa2XMP"
nDos = 0

' S�lection du dossier contenant le fichier .picasa.ini
tDossier = tSelect_Dossier( "" )

' Tests sur le dossier choisi
If tDossier = vbNull Then
    WScript.Echo "Annul�"
Else
	' Parcours r�cursif de l'arobrescence s�lectionn�e et traitement de chaque fichier ".picasa.ini" trouv�
	Set FSO = CreateObject("Scripting.FileSystemObject")
	Traite_dossier tDossier
End If

ProgressMsg "Termin�", tWindowTitle
WScript.Quit

'----------------------------------------------------------------------

Function tSelect_Dossier( myStartFolder )
' This function opens a "Select Folder" dialog and will
' return the fully qualified path of the selected folder
'
' Argument:
'     myStartFolder    [string]    the root folder where you can start browsing;
'                                  if an empty string is used, browsing starts
'                                  on the local computer
'
' Returns:
' A string containing the fully qualified path of the selected folder
'
' Written by Rob van der Woude
' http://www.robvanderwoude.com

    ' Standard housekeeping
    Dim objFolder, objItem, objShell
    
    ' Custom error handling
    On Error Resume Next
    tSelect_Dossier = vbNull

    ' Create a dialog object
    Set objShell  = CreateObject( "Shell.Application" )
    Set objFolder = objShell.BrowseForFolder( 0, "Dossier de d�part", 0, myStartFolder )

    ' Return the path of the selected folder
    If IsObject( objfolder ) Then tSelect_Dossier = objFolder.Self.Path

    ' Standard housekeeping
    Set objFolder = Nothing
    Set objshell  = Nothing
    On Error Goto 0
End Function


Sub Traite_dossier(path)
' Ce sub scrute le dossier en cours � la recherche du fichier .picasa.ini
' Il effectue 2 t�ches :
' 1 - Il traite le fichier .picasa.ini trouv�
' 2 - Il se lance de mani�re r�cursive pour traiter les sous-dossiers trouv�s dans le dossier en cours

  Dim folder, files, file, subfolders,subfolder
  Dim FSP
  
  set folder = FSO.getFolder(path)
  set files = folder.Files
  Set FSP = CreateObject("Scripting.FileSystemObject")
  
  ' Traitement du dossier trouv� � condition qu'il contienne un fichier .picasa.ini
  If FSP.FileExists(FSP.BuildPath(path, ".picasa.ini")) Then
    Traite_picasa path
  End If

  set subFolders = folder.subFolders
  
  For Each subfolder in subFolders
    Traite_dossier subfolder.path
  Next

End Sub

'----------------------------------------------------------------------

Sub Traite_picasa(pPath)
' Ce sub effectue le traitement du fichier .picasa.ini qui vient d'�tre trouv�
' Le .picasa.ini est lu en 2 passes :
'
' Passe 1 :
' - Les lignes commen�ant par [.album:.... sont localis�es et l'Id de l'album est stock� dans un tableau
' - La propri�t� "name=" qui se trouve sur la ligne qui suit est m�moris�e dans un autre tableau avec le m�me indice
'
' Le fichier "albums.tsv" est lu. Chaque ligne est compar�e � une propri�t� "name" trouv�e dans le dossier en cours. Si une correspondance est trouv�e le keyword XMP correspondant est m�moris� dans un tableau avec le m�me indice
'
' Passe 2 :
' - Les lignes [xxxx.jpg] sont localis�es
' La propri�t� "albums=" qui se trouve sur les lignes suivantes est lue. Si elle correspond � l'un des Id d'albums m�moris�s alors on lance la mise � jour du fichier XMP pour cette photo

  Dim FSI, FSL
  Dim tLigne, tId, tId_liste, tAlb, tKey, tPhoto
  Dim nBalbums
  Dim nCur, nPtr, nStar
  
  Dim aId(128), aName(128), aKeyword(128)

  Dim WshShell, tCurdir

  'WScript.echo "Traitement du dossier " + pPath
  ProgressMsg "Traitement de " + pPath, tWindowTitle
  
  ' R�cup�ration du r�pertoire en cours
  Set WshShell = CreateObject("WScript.Shell")
  tCurdir = WshShell.CurrentDirectory
  
  'WScript.echo "Dossier courant : " + tCurdir
  
  ' Ouverture du fichier LOG
  nDos = nDos + 1
  Set FSL = CreateObject("ADODB.Stream")
  FSL.Open
  FSL.Charset = "utf-8"
  FSL.Type    = 2
    
 
  nBalbums = 0
  
  FSL.WriteText "Traitement du dossier " + pPath + vbCrLf
    
  
  ' PASSE 1
  '
  ' recherche des lignes [.album:, puis m�morisation de l'Id pr�sent sur la ligne "name=" qui suit

  ' Ouverture d'un Stream, pour la lecture de .picasa.ini qui est encod� en UTF-8
  Set FSI = CreateObject("ADODB.Stream")
  FSI.Open
  FSI.Charset = "utf-8"
  FSI.Type    = 2
  FSI.LoadFromFile pPath + "\.picasa.ini"

  nCur = 0
  tId = ""
  
  Do Until FSI.EOS
    tLigne = FSI.ReadText(-2)
	
	If left(tLigne, 7) = "[.album" Then
	  ' On lit l'Id de l'album dans une variable temporaire, en attendant de trouver la propri�t� "name=" correspondante
	  tId = Mid(tLigne, 9, 32)
	End If
	
	If left(tLigne, 5) = "name=" and tId <> "" Then
	  ' On a trouv� une propri�t� "name=", il faut m�moriser l'Id et le nom de l'album dans les tableaux
	  nCur = nCur + 1
	  aId(nCur) = tId
	  aName(nCur) = Mid(tLigne, 6)
	  tId = ""
	  
      FSL.WriteText "Album pr�sent : " + cstr(nCur) + " : " + aName(nCur) + " - " + aId(nCur) + vbCrLf
	End If
  Loop
  
  ' M�morisation du nombre d'albums pour ce dossier
  nBalbums = nCur
  FSL.WriteText Cstr(nBalbums) + " albums pr�sents" + vbCrLf
  
  FSI.Close

  
  '
  ' Lecture du fichier "albums.tsv"
  '
  ' recherche les correspondances XMP aux albums Picasa trouv�s
  
  FSI.Open
  FSI.Charset = "utf-8"
  FSI.Type    = 2
  FSI.LoadFromFile tCurdir + "\albums.tsv"

  ' On se d�barasse de la 1�re ligne qui contient les ent�tes
  tLigne = FSI.ReadText(-2)
  
  Do Until FSI.EOS
    tLigne = FSI.ReadText(-2)
	
	' Localisation du s�parateur et chargement du nom de l'album
	nPtr = instr(tLigne, chr(9))
	tAlb = left(tLigne, nPtr - 1)
	tKey = mid(tLigne, nPtr + 1)
	
	'WScript.echo "Album.tsv : " + tAlb + "---" + tKey
	
	' Recherche d'une correspondance dans les albums du dossier en cours
    For nCur = 1 To nBalbums
	  If aName(nCur) = tAlb Then
        aKeyword(nCur) = tKey
        
		FSL.WriteText "Correspondance XMP trouv�e : " + aName(nCur) + " --> " + aKeyword(nCur) + vbCrLf
	  End If
	Next
  Loop

  FSI.Close
  
  FSL.WriteText vbCrLf + "-----------------------------------------" + vbCrLf
  FSL.WriteText "Lecture du fichier .picasa.ini pour traitement des photos" + vbCrLf + vbCrLf
  
    
  '
  ' PASSE 2
  '
  ' Nouvelle lecture de .picasa.ini � la recherche des photos [xxx.jpg], puis des rattachements aux albums "albums=", et du rating "star=yes"
  
  FSI.Open
  FSI.Charset = "utf-8"
  FSI.Type    = 2
  FSI.LoadFromFile pPath + "\.picasa.ini"

  Do Until FSI.EOS
    tLigne = FSI.ReadText(-2)
	
    ' Ligne PHOTO
	If Right(tLigne, 5) = ".jpg]"_ 
	or Right(tLigne, 5) = ".JPG]"_
	or Right(tLigne, 5) = ".png]"_
	or Right(tLigne, 5) = ".PNG]"_
	or Right(tLigne, 5) = ".gif]"_
	or Right(tLigne, 5) = ".GIF]"_
	or Right(tLigne, 5) = ".TIF]"_
	or Right(tLigne, 5) = ".tif]" Then
	  ' On lit le nom de la photo dans une variable temporaire, en attendant de trouver la propri�t� "albums=" �ventuelle
	  tPhoto = Mid(tLigne, 2, Len(tLigne) - 2)
	  nStar = 0

      'WScript.echo "Ligne (photo) : " + tLigne + "(" + tPhoto + ")"
	End If
	
	' Ligne RATING
	If tLigne = "star=yes" Then
	  nStar = 3
	End If
	
	' Ligne ALBUM
	If left(tLigne, 7) = "albums=" Then
	  ' On a trouv� une propri�t� "albums", il faut scruter le ou les Id qui suivent et comparer avec ceux du tableau aId
	  'Wscript.echo "Ligne albums : " + tLigne
	  
	  ' La liste des Id commence apr�s "albums"
	  tId_liste = Mid(tLigne, 8)
	  tKey = ""
	  ' Chaque Id fait 32 octets de long
	  tId = Mid(tId_liste, 1, 32)
	  
	  Do Until tId = ""
	    'Wscript.echo "Boucle tId / tId_liste : " + tId + " / " + tId_liste

	  ' Recherche d'une correspondance dans les albums du dossier en cours
        For nCur = 1 To nBalbums
	      If aId(nCur) = tId Then
            'WScript.echo "Affectation trouv�e : " + tPhoto + " - " + aKeyword(nCur)
			
            ' On v�rifie que l'affectation trouv�e n'est pas d�j� pr�sente dans les Keywords en cours
			nPtr = Instr(tKey, aKeyword(nCur))
			
			'WScript.echo "Photo " + tPhoto + "(" + tKey + ") + " + aKeyword(nCur) + " nPtr = " + cStr(nPtr)
			
			If nPtr = 0 or tKey = "" Then
			  ' Le Keyword trouv� est nouveau, on l'ajoute
			  If tKey <> "" Then
			    tKey = tKey + "," + aKeyword(nCur)
			    'WScript.echo "Keywords : " + tKey
              Else
			    tKey = aKeyword(nCur)
			  End If
			End If
	      End If
	    Next

		' On passe sur l'Id suivant dans la liste
		tId_liste = Mid(tId_liste, 34)
	    tId = Mid(tId_liste, 1, 32)
	  Loop
	  
	  ' Lancement de l'�criture du XMP
      If tKey <> "" or nStar > 0 Then
        FSL.WriteText "Affectation trouv�e : " + tPhoto + " - " + tKey + " (" + Cstr(nStar) + " �toile(s))" + vbCrLf
	    MAJ_XMP pPath + "\" + tPhoto, tKey, nStar
	  End If
	  
	End If
  Loop
  
  FSI.Close

  ' Fermeture du fichier LOG
  FSL.WriteText vbCrLf + "-----------------------------------------" + vbCrLf
  FSL.WriteText "Fin du traitement de " + pPath + vbCrLf + vbCrLf
  FSL.SaveToFile tCurdir + "\picasa2xmp_" + Cstr(nDos) + ".log", 2
  FSL.close  
  
  
  ' Lib�ration des allocations de m�moire
  Erase aId
  Erase aName
  Erase aKeyword
  
End Sub

'----------------------------------------------------------------------

Sub MAJ_XMP(ficJPG, tKeyword, nRating)
' Ce sub met � jour le fichier XMP qui correspond � la photo 'ficJPG' et y associe le mot-cl� 'tKeyword' ainsi que le rating

  Dim FSR, FSW, FSS
  Dim tIn, tOut1, tOut2
  Dim nPtr

  ' Ouverture d'un Stream, pour la lecture du fichier XMP qui est encod� en UTF-8
  Set FSR = CreateObject("ADODB.Stream")
  FSR.Open
  FSR.Charset = "utf-8"
  FSR.Type    = 2
  FSR.LoadFromFile ficJPG + ".xmp"
  
  ' Lecture de la totalit� du fichier XMP
  tIn = FSR.ReadText()
  FSR.close

  ' Gestion du Keyword
  If tKeyword <> "" Then
    ' Localisation de la propri�t� qui nous int�resse : keywordlist. Elle doit �tre nulle
    nPtr = Instr(tIn, "bopt:keywordlist=""""")
    
    ' Si la propri�t� keywordlist est trouv�e, on r��crit le fichier
    If nPtr > 0 Then
	  ' On ins�re le Keyword dans la sortie
	  tOut1 = Left(tIn, nPtr - 1) + "bopt:keywordlist=""" + tKeyword + """" + Mid(tIn, nPtr + 19)
	Else
	  tOut1 = tIn
	End If
  Else
    tOut1 = tIn
  End If
  
  ' Gestion du rating
  If nRating > 0 Then
    ' Localisation de la propri�t� qui nous int�resse : rating. Elle doit �tre nulle
    nPtr = Instr(tOut1, "bopt:rating=""0""")
    
    ' Si la propri�t� rating est trouv�e, on r��crit le fichier
    If nPtr > 0 Then
	  ' On ins�re le Keyword dans la sortie
	  tOut2 = Left(tOut1, nPtr - 1) + "bopt:rating=""" + Cstr(nRating) + """" + Mid(tOut1, nPtr + 15)
	Else
	  tOut2 = tOut1
	End If
  Else
    tOut2 = tOut1
  End If
  
  ' Ouverture d'un Stream, pour l'�criture du nouveau fichier XMP encod� en UTF-8
  Set FSW = CreateObject("ADODB.Stream")
  FSW.Open
  FSW.Charset = "utf-8"
  FSW.Type    = 2
    
  FSW.WriteText tOut2
    
  FSW.SaveToFile ficJPG + ".xmp_tmp", 2
  FSW.close
    
  ' Remplacement de l'ancien fichier XMP par le nouveau
  Set FSS = CreateObject("Scripting.FileSystemObject")
  FSS.DeleteFile ficJPG + ".xmp"
  FSS.MoveFile ficJPG + ".xmp_tmp", ficJPG + ".xmp"
  
End Sub

Function ProgressMsg( strMessage, strWindowTitle )
' Written by Denis St-Pierre
' Displays a progress message box that the originating script can kill in both 2k and XP
' If StrMessage is blank, take down previous progress message box
' Using 4096 in Msgbox below makes the progress message float on top of things
' CAVEAT: You must have   Dim ObjProgressMsg   at the top of your script for this to work as described
'
' Modif : on stocke le timestamp de la derni�re ex�cution pour �viter 2 appels trop rapproch�s
'
    Dim wshShell, objFSO, objTempMessage
	Dim strTEMP, strTempVBS
	
	If nTimer = Timer Then
	  WScript.Sleep 1000
	End If
	
	Set wshShell = CreateObject( "WScript.Shell" )
    strTEMP = wshShell.ExpandEnvironmentStrings( "%TEMP%" )
    If strMessage = "" Then
        ' Disable Error Checking in case objProgressMsg doesn't exists yet
        On Error Resume Next
        ' Kill ProgressMsg
        objProgressMsg.Terminate( )
        ' Re-enable Error Checking
        On Error Goto 0
        Exit Function
    End If
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    strTempVBS = strTEMP + "\" & "Message.vbs"     'Control File for reboot

    ' Create Message.vbs, True=overwrite
    Set objTempMessage = objFSO.CreateTextFile( strTempVBS, True )
    objTempMessage.WriteLine( "MsgBox""" & strMessage & """, 4096, """ & strWindowTitle & """" )
    objTempMessage.Close

    ' Disable Error Checking in case objProgressMsg doesn't exists yet
    On Error Resume Next
    ' Kills the Previous ProgressMsg
    objProgressMsg.Terminate( )
    ' Re-enable Error Checking
    On Error Goto 0

    ' Trigger objProgressMsg and keep an object on it
    Set objProgressMsg = WshShell.Exec( "%windir%\system32\wscript.exe " & strTempVBS )
	
	' M�morisation du temps � la derni�re ex�cution, pour �viter 2 appels trop rapproch�s
	nTimer = Timer

    Set wshShell = Nothing
    Set objFSO   = Nothing
End Function
