Attribute VB_Name = "mod_Main"
'====================================================================================================================
' Version: 1.0
' Datum: 19.04.2022
' Autor: M. Shtekhno

'--------------------------------------------------
' Beschreibung:
'--------------------------------------------------
' • Anleitung finden Sie am Ende des Moduls!
' •
' •

'--------------------------------------------------
' Update Log: 1.0 -> 1.1
'--------------------------------------------------
' •

'--------------------------------------------------
' Update Log: 1.1 -> 1.2
'--------------------------------------------------
' • ...

'====================================================================================================================

'TODO:
' •
' •

'Option Explicit

'********************************************************************************************************************
'Klassenvariablen
'********************************************************************************************************************

Private maxiTools As cls_MaxiTools 'Klasse mit den hilfreichen Methoden
Private maxiBar As cls_MaxiBar 'Klasse mit den hilfreichen Methoden


'********************************************************************************************************************
'Mainmethoden

'Beschreibung : Main Methode. Dies ist das Modul, das ausgeführt wird, wenn die Schaltfläche im Excel-Fenster aktiviert wird.
'Verantwortlich für die Ausführung der Programmschritte in der richtigen Reihenfolge und die Überwachung des Erfolgs der Schritte
'********************************************************************************************************************

'*************************
'Beschreibung : Main Beschreibung

'Args: -
'Returns: -
'**************************
Sub main()

    On Error GoTo errorExit 'Wenn ein Fehler im Code auftritt, wird die Codeausführung gestoppt und das errorExit-Modul automatisch gestartet

    Set maxiTools = New cls_MaxiTools 'MaxiTools definieren
    Set maxiBar = New cls_MaxiBar 'MaxiTools definieren
    
    maxiTools.disableAppSettings 'Excel Einstellungen ausschalten -> Beschleunigung
        
        maxiBar.openStatusBar 'Starten einer Fortschrittsleiste
        
        maxiBar.runStatusBar 1, 2, "First out of five..."

    maxiTools.enableAppSettings 'Excel Einstellungen wieder einschlten
    
    maxiBar.deleteBar 'Beendigung der Fortschrittsleiste nach Beendigung des Makros
    
    MsgBox "Fertig!", vbInformation, "Fertig!" 'Benachrichtigung des Benutzers, wenn ein Makro erfolgreich abgeschlossen wurde
    
    Exit Sub

'Dieser Block wird ausgeführt, wenn ein Fehler im Main Modul auftritt
errorExit:
    
    MsgBox errorDescription, vbCritical, "Fehler!" 'Generierung einer Nutzermeldung ohne technische Details und mit vordefiniertem Text
    Debug.Print Err.Number & Err.Description & Err.Source 'Erstellung einer technischen Mitteilung für Entwickler mit technischen Details und vordefinierten Fehlerparametern
    maxiTools.enableAppSettings 'Aktivieren von Excel-Einstellungen im Falle eines Fehlers
    maxiBar.deleteBar 'Ausschalten der Fortschrittsleiste im Falle eines Fehlers
    Application.CutCopyMode = False 'Selection löschen
    
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
    errorDescription = "Fehler in der Sunmethode! Bitte prüfen Sie die technische Beschreibung des Fehlers oder wenden Sie sich an einen zuständigen Support-Mitarbeiter!"

    maxiBar.runStatusBar 2, 2, "Submethode Nummer 1..."

    
    
    Exit Sub
    
'Erzeugung eines benutzerdefinierten Fehlers, der dazu führt, dass das Makro die Ausführung beendet und das errorExit-Modul in Main ausführt
errorExit:
    Err.Raise 2012, "subMethode", "Mögliche Ursachen des Fehlers: ..."
    
End Sub



'********************************************************************************************************************
'Hilfsmethoden
'********************************************************************************************************************

'*************************
'Beschreibung : Dieses Modul führt die erste Kommunikation mit dem Benutzer durch, und wenn der Benutzer sich weigert
'das Makro auszuführen, wird das Programm beendet.

'Args: -
'Returns: -
'*************************
Sub startUserDialog()

    errorDescription = "Der Benutzer hat die Ausführung des Makros unterbrochen!" 'Angabe einer Fehlerbeschreibung für den Fall, dass ein Fehler in diesem Modul auftritt

    Dim userAnswer As Integer 'Variable zum Speichern der Entscheidung des Users
    
    userAnswer = MsgBox("Makro gestartet. Fortfahren?", vbQuestion + vbYesNo + vbDefaultButton1, "Makro gestartet") 'Die Frage für den Benutzer ist, ob er das Makro weiter ausführen möchte. Antwortmöglichkeiten: Ja/Nein
    
    'Wenn der Benutzer das Makro nicht ausführen will, wird ein Fehler erzeugt, der zum Abbruch des Programms führt.
    If userAnswer = vbNo Then
    
        Err.Raise 2012, "startUserDialog", "Der Benutzer hat die Ausführung des Makros am Anfang der Frage in msgbox unterbrochen"
    
    End If

End Sub



'********************************************************************************************************************
'Universelle Vorlage - Beschreibung und Anwendungsbeispiel
'********************************************************************************************************************

'Dieses Modul ist eine Vorlage für alle Arten und Formate von Projekten. Wenn Sie sich an diese Vorlage halten, können Sie in kurzer Zeit
'ein Projekt realisieren, das zunächst solche Dinge implementiert wie: Fortschrittsbalken, Fehlerabfangmodul, Klasse mit meinen persönlichen Methoden.
'
'Der Code wird im Modul mod_Main implementiert, das bereits eine komplette Landschaft zum Schreiben von Code im Voraus hat.
'In der Kopfzeile des Moduls lohnt es sich, auf die kompetente Vervollständigung der Modulbeschreibung und seiner Version zu achten, und eine sorgfältige
'Pflege des UpdateLogs wird empfohlen, um den Entwicklungsfortschritt des Projekts besser nachvollziehen zu können. Die ToDo-Liste ist optional, ermöglicht
'aber ein besseres Verständnis der Dinge, die mit den kommenden Updates implementiert werden sollten.
'
'Im ersten Modul - Klassenvariablen - lohnt es sich, diejenigen Variablen zu initialisieren, die in mehreren Modulen oder Untermodulen auf einmal verwendet
'werden sollen. Wenn eines der Untermodule den Wert einer bestimmten Variablen ändert, wird dies für alle anderen Module sichtbar. Dies ist eine Implementierung
'von Funktionen, die der statischen Funktion in Java ähneln. Auch in diesem Teil des Codes müssen Sie die Variablen initialisieren, die in Zukunft den Wert der
'Instanzen der Klassen Tools und MaxiBar annehmen werden.
'
'Mainmethoden enthalten eine oder mehrere Hauptmethoden, die für den Ablauf des Moduls selbst und die Reihenfolge der Ausführung der Untermodule verantwortlich
'sind und das Hauptmodul zur Fehlerbehandlung enthalten. Hier werden auch die Module des Fortschrittsbalkens behandelt. Das Modul Fortschrittsbalken wird im Modul selbst erläutert.
'
'Eine kurze Beschreibung der Fehlerbehandlung und des Algorithmus zur Fehlerbehandlung.
'In VBA gibt es keine Funktionen zur Behandlung von Ausnahmen (wie Try/Catch in anderen Programmiersprachen). Um Funktionen zur Behandlung von Ausnahmen zu implementieren,
'muss man also zusätzliche Tools verwenden und sie selbst erstellen. Die Grundidee der Fehlerbehandlung in diesem Modul:
'
'1) Am Anfang jedes Moduls gibt es eine Zeile "on Error goto errorExit". Diese Zeile stellt sicher, dass das Modul bei Auftreten eines system- oder benutzergenerierten Fehlers
'automatisch gestoppt und der exitError-Block gestartet wird. Besonders hervorzuheben ist in diesem Zusammenhang die Fehlerbehandlung im Modul Main. Er ist für den Betrieb aller
'Teilmodule verantwortlich und überwacht daher Fehler in allen Teilmodulen. Tritt in einem beliebigen Submodul ein Fehler auf, so wird zuerst der errorExit des Submoduls ausgelöst
'und dann sofort der errorExit des Hauptmoduls.
'Daher wird im errorExit im Submodul nur die Variable mit der Fehlerbeschreibung initialisiert, und der Makro-Beendigungscode wird einmal für alle Fälle im Hauptmodul geschrieben.
'
'2) ErrorExit im Hauptmodul ist, wie oben erwähnt, das Hauptmodul und wird im Falle eines Fehlers in einem beliebigen Modul ausgeführt. Sie ist anders als alle anderen.
'Es gibt gleich mehrere wichtige Maßnahmen zur Erhöhung der Benutzerfreundlichkeit. Zunächst wird eine Meldung an den Benutzer auf der Grundlage einer vorher festgelegten
'Fehlerquelle generiert. Die Fehlerbeschreibung ist eine Variable, deren Wert manuell eingestellt wird, damit ihr Text für den Benutzer so klar und einfach wie möglich ist.
'Auch hier wird die Meldung für die Entwickler generiert, die im Debugger angezeigt wird und technischen Charakter hat - die Fehlernummer, die technische Beschreibung und
'der Name des Moduls, in dem der Fehler aufgetreten ist, wobei die Meldung möglicherweise durch zusätzliche Parameter erweitert wird. Für künstlich erzeugte Fehler verwende
'ich die Fehlernummer 2012. Diese Zahl ist frei wählbar und kann vom Entwickler selbst festgelegt werden, um die Art des Fehlers besser zu verstehen. 2012 ist meine Glückszahl,
'also habe ich sie für die Zahl meiner eigenen Fehler gewählt. Auch nach der Erzeugung des Fehlers und der Fertigstellung des Makros müssen die Excel-Einstellungen wieder verbunden
'und der Fortschrittsbalken entfernt werden.
'
'3) ErrorExit wird in allen Submodulen für denselben Zweck verwendet - zur Erzeugung eines Benutzerfehlers mit der Methode Err.Raise (ähnlich wie throw Exception in anderen Programmiersprachen)
'und zur Angabe der Benutzer- und technischen Beschreibung dieses Fehlers. Manchmal gibt es in Modulen explizite Fälle, in denen eine manuelle Fehlergenerierung verwendet werden sollte,
'aber in den meisten Modulen erfolgt die Überprüfung im Hintergrund.
'
'Die Untermodule sind das zweitwichtigste Element des Programms. Sie sollten die einzelnen Schritte seiner Ausführung beschreiben und diese dann in der richtigen Reihenfolge im
'Hauptmodul selbst aufrufen. Jedes Submodul sollte aus einer modulspezifischen Fehlerbeschreibung, einer Methode zur Aktualisierung des Fortschrittsbalkens und errorExit für den Fall,
'dass ein Fehler auftritt, bestehen.
'
'Hilfsmethoden unterscheiden sich nur dadurch, dass ihr Vorhandensein für den Verlauf des Programms selbst nicht so wichtig ist, oder sie sind Teil von Untermodulen, die zu einer
'separaten Funktion ausgelagert werden. Eine davon ist die Standardeinstellung - die Startansage des Benutzers.
'
'Jede Methode muss durch ihre Beschreibung und, gemäß dem Google-Standard, durch eine Beschreibung ihrer Parameter und Rückgaben dokumentiert werden.






