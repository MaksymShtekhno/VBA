Attribute VB_Name = "mod_Main"
'====================================================================================================================
' Version: 1.0
' Datum: 19.04.2022
' Autor: M. Shtekhno

'--------------------------------------------------
' Beschreibung:
'--------------------------------------------------
' � Anleitung finden Sie am Ende des Moduls!
' �
' �

'--------------------------------------------------
' Update Log: 1.0 -> 1.1
'--------------------------------------------------
' �

'--------------------------------------------------
' Update Log: 1.1 -> 1.2
'--------------------------------------------------
' � ...

'====================================================================================================================

'TODO:
' �
' �

'Option Explicit

'********************************************************************************************************************
'Klassenvariablen
'********************************************************************************************************************

Private maxiTools As cls_MaxiTools 'Klasse mit den hilfreichen Methoden
Private maxiBar As cls_MaxiBar 'Klasse mit den hilfreichen Methoden


'********************************************************************************************************************
'Mainmethoden

'Beschreibung : Main Methode. Dies ist das Modul, das ausgef�hrt wird, wenn die Schaltfl�che im Excel-Fenster aktiviert wird.
'Verantwortlich f�r die Ausf�hrung der Programmschritte in der richtigen Reihenfolge und die �berwachung des Erfolgs der Schritte
'********************************************************************************************************************

'*************************
'Beschreibung : Main Beschreibung

'Args: -
'Returns: -
'**************************
Sub main()

    On Error GoTo errorExit 'Wenn ein Fehler im Code auftritt, wird die Codeausf�hrung gestoppt und das errorExit-Modul automatisch gestartet

    Set maxiTools = New cls_MaxiTools 'MaxiTools definieren
    Set maxiBar = New cls_MaxiBar 'MaxiTools definieren
    
    maxiTools.disableAppSettings 'Excel Einstellungen ausschalten -> Beschleunigung
        
        maxiBar.openStatusBar 'Starten einer Fortschrittsleiste
        
        maxiBar.runStatusBar 1, 2, "First out of five..."

    maxiTools.enableAppSettings 'Excel Einstellungen wieder einschlten
    
    maxiBar.deleteBar 'Beendigung der Fortschrittsleiste nach Beendigung des Makros
    
    MsgBox "Fertig!", vbInformation, "Fertig!" 'Benachrichtigung des Benutzers, wenn ein Makro erfolgreich abgeschlossen wurde
    
    Exit Sub

'Dieser Block wird ausgef�hrt, wenn ein Fehler im Main Modul auftritt
errorExit:
    
    MsgBox errorDescription, vbCritical, "Fehler!" 'Generierung einer Nutzermeldung ohne technische Details und mit vordefiniertem Text
    Debug.Print Err.Number & Err.Description & Err.Source 'Erstellung einer technischen Mitteilung f�r Entwickler mit technischen Details und vordefinierten Fehlerparametern
    maxiTools.enableAppSettings 'Aktivieren von Excel-Einstellungen im Falle eines Fehlers
    maxiBar.deleteBar 'Ausschalten der Fortschrittsleiste im Falle eines Fehlers
    Application.CutCopyMode = False 'Selection l�schen
    
End Sub

'********************************************************************************************************************
'Submethoden
'********************************************************************************************************************


'*************************
'Beschreibung :

'Args: -
'Returns: -
'*************************

Sub subMethode()

    On Error GoTo errorExit
    errorDescription = "Fehler in der Sunmethode! Bitte pr�fen Sie die technische Beschreibung des Fehlers oder wenden Sie sich an einen zust�ndigen Support-Mitarbeiter!"

    maxiBar.runStatusBar 2, 2, "Submethode Nummer 1..."

    
    
    Exit Sub
    
'Erzeugung eines benutzerdefinierten Fehlers, der dazu f�hrt, dass das Makro die Ausf�hrung beendet und das errorExit-Modul in Main ausf�hrt
errorExit:
    Err.Raise 2012, "subMethode", "M�gliche Ursachen des Fehlers: ..."
    
End Sub



'********************************************************************************************************************
'Hilfsmethoden
'********************************************************************************************************************

'*************************
'Beschreibung : Dieses Modul f�hrt die erste Kommunikation mit dem Benutzer durch, und wenn der Benutzer sich weigert
'das Makro auszuf�hren, wird das Programm beendet.

'Args: -
'Returns: -
'*************************
Sub startUserDialog()

    errorDescription = "Der Benutzer hat die Ausf�hrung des Makros unterbrochen!" 'Angabe einer Fehlerbeschreibung f�r den Fall, dass ein Fehler in diesem Modul auftritt

    Dim userAnswer As Integer 'Variable zum Speichern der Entscheidung des Users
    
    userAnswer = MsgBox("Makro gestartet. Fortfahren?", vbQuestion + vbYesNo + vbDefaultButton1, "Makro gestartet") 'Die Frage f�r den Benutzer ist, ob er das Makro weiter ausf�hren m�chte. Antwortm�glichkeiten: Ja/Nein
    
    'Wenn der Benutzer das Makro nicht ausf�hren will, wird ein Fehler erzeugt, der zum Abbruch des Programms f�hrt.
    If userAnswer = vbNo Then
    
        Err.Raise 2012, "startUserDialog", "Der Benutzer hat die Ausf�hrung des Makros am Anfang der Frage in msgbox unterbrochen"
    
    End If

End Sub



'********************************************************************************************************************
'Universelle Vorlage - Beschreibung und Anwendungsbeispiel
'********************************************************************************************************************

'Dieses Modul ist eine Vorlage f�r alle Arten und Formate von Projekten. Wenn Sie sich an diese Vorlage halten, k�nnen Sie in kurzer Zeit
'ein Projekt realisieren, das zun�chst solche Dinge implementiert wie: Fortschrittsbalken, Fehlerabfangmodul, Klasse mit meinen pers�nlichen Methoden.
'
'Der Code wird im Modul mod_Main implementiert, das bereits eine komplette Landschaft zum Schreiben von Code im Voraus hat.
'In der Kopfzeile des Moduls lohnt es sich, auf die kompetente Vervollst�ndigung der Modulbeschreibung und seiner Version zu achten, und eine sorgf�ltige
'Pflege des UpdateLogs wird empfohlen, um den Entwicklungsfortschritt des Projekts besser nachvollziehen zu k�nnen. Die ToDo-Liste ist optional, erm�glicht
'aber ein besseres Verst�ndnis der Dinge, die mit den kommenden Updates implementiert werden sollten.
'
'Im ersten Modul - Klassenvariablen - lohnt es sich, diejenigen Variablen zu initialisieren, die in mehreren Modulen oder Untermodulen auf einmal verwendet
'werden sollen. Wenn eines der Untermodule den Wert einer bestimmten Variablen �ndert, wird dies f�r alle anderen Module sichtbar. Dies ist eine Implementierung
'von Funktionen, die der statischen Funktion in Java �hneln. Auch in diesem Teil des Codes m�ssen Sie die Variablen initialisieren, die in Zukunft den Wert der
'Instanzen der Klassen Tools und MaxiBar annehmen werden.
'
'Mainmethoden enthalten eine oder mehrere Hauptmethoden, die f�r den Ablauf des Moduls selbst und die Reihenfolge der Ausf�hrung der Untermodule verantwortlich
'sind und das Hauptmodul zur Fehlerbehandlung enthalten. Hier werden auch die Module des Fortschrittsbalkens behandelt. Das Modul Fortschrittsbalken wird im Modul selbst erl�utert.
'
'Eine kurze Beschreibung der Fehlerbehandlung und des Algorithmus zur Fehlerbehandlung.
'In VBA gibt es keine Funktionen zur Behandlung von Ausnahmen (wie Try/Catch in anderen Programmiersprachen). Um Funktionen zur Behandlung von Ausnahmen zu implementieren,
'muss man also zus�tzliche Tools verwenden und sie selbst erstellen. Die Grundidee der Fehlerbehandlung in diesem Modul:
'
'1) Am Anfang jedes Moduls gibt es eine Zeile "on Error goto errorExit". Diese Zeile stellt sicher, dass das Modul bei Auftreten eines system- oder benutzergenerierten Fehlers
'automatisch gestoppt und der exitError-Block gestartet wird. Besonders hervorzuheben ist in diesem Zusammenhang die Fehlerbehandlung im Modul Main. Er ist f�r den Betrieb aller
'Teilmodule verantwortlich und �berwacht daher Fehler in allen Teilmodulen. Tritt in einem beliebigen Submodul ein Fehler auf, so wird zuerst der errorExit des Submoduls ausgel�st
'und dann sofort der errorExit des Hauptmoduls.
'Daher wird im errorExit im Submodul nur die Variable mit der Fehlerbeschreibung initialisiert, und der Makro-Beendigungscode wird einmal f�r alle F�lle im Hauptmodul geschrieben.
'
'2) ErrorExit im Hauptmodul ist, wie oben erw�hnt, das Hauptmodul und wird im Falle eines Fehlers in einem beliebigen Modul ausgef�hrt. Sie ist anders als alle anderen.
'Es gibt gleich mehrere wichtige Ma�nahmen zur Erh�hung der Benutzerfreundlichkeit. Zun�chst wird eine Meldung an den Benutzer auf der Grundlage einer vorher festgelegten
'Fehlerquelle generiert. Die Fehlerbeschreibung ist eine Variable, deren Wert manuell eingestellt wird, damit ihr Text f�r den Benutzer so klar und einfach wie m�glich ist.
'Auch hier wird die Meldung f�r die Entwickler generiert, die im Debugger angezeigt wird und technischen Charakter hat - die Fehlernummer, die technische Beschreibung und
'der Name des Moduls, in dem der Fehler aufgetreten ist, wobei die Meldung m�glicherweise durch zus�tzliche Parameter erweitert wird. F�r k�nstlich erzeugte Fehler verwende
'ich die Fehlernummer 2012. Diese Zahl ist frei w�hlbar und kann vom Entwickler selbst festgelegt werden, um die Art des Fehlers besser zu verstehen. 2012 ist meine Gl�ckszahl,
'also habe ich sie f�r die Zahl meiner eigenen Fehler gew�hlt. Auch nach der Erzeugung des Fehlers und der Fertigstellung des Makros m�ssen die Excel-Einstellungen wieder verbunden
'und der Fortschrittsbalken entfernt werden.
'
'3) ErrorExit wird in allen Submodulen f�r denselben Zweck verwendet - zur Erzeugung eines Benutzerfehlers mit der Methode Err.Raise (�hnlich wie throw Exception in anderen Programmiersprachen)
'und zur Angabe der Benutzer- und technischen Beschreibung dieses Fehlers. Manchmal gibt es in Modulen explizite F�lle, in denen eine manuelle Fehlergenerierung verwendet werden sollte,
'aber in den meisten Modulen erfolgt die �berpr�fung im Hintergrund.
'
'Die Untermodule sind das zweitwichtigste Element des Programms. Sie sollten die einzelnen Schritte seiner Ausf�hrung beschreiben und diese dann in der richtigen Reihenfolge im
'Hauptmodul selbst aufrufen. Jedes Submodul sollte aus einer modulspezifischen Fehlerbeschreibung, einer Methode zur Aktualisierung des Fortschrittsbalkens und errorExit f�r den Fall,
'dass ein Fehler auftritt, bestehen.
'
'Hilfsmethoden unterscheiden sich nur dadurch, dass ihr Vorhandensein f�r den Verlauf des Programms selbst nicht so wichtig ist, oder sie sind Teil von Untermodulen, die zu einer
'separaten Funktion ausgelagert werden. Eine davon ist die Standardeinstellung - die Startansage des Benutzers.
'
'Jede Methode muss durch ihre Beschreibung und, gem�� dem Google-Standard, durch eine Beschreibung ihrer Parameter und R�ckgaben dokumentiert werden.






