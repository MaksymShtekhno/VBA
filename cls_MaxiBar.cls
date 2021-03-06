VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_MaxiBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'====================================================================================================================
' Version: 1.0
' Datum: 12.04.2022
' Autor: M. Shtekhno
'--------------------------------------------------
' Beschreibung:
'--------------------------------------------------
' � Die Klasse, die f�r die Verwaltung des Fortschrittsbalkens w�hrend der Programmausf�hrung zust�ndig ist. Funktioniert in Verbindung mit dem Formular MaxiBar, das ebenfalls im Projekt vorhanden sein muss.
' � Beschreibung des Formulars:
'      � Breite des Fortschrittsbalkens: 264
'      � Anleitung finden Sie am Ende des Moduls!
'--------------------------------------------------
' Update Log: 1.0 -> 1.1
'--------------------------------------------------
' � Beschrebung :
'====================================================================================================================

'*************************
'Beschreibung : Diese Methode ist f�r die Anzeige und die Ersteinstellung des Fortschrittsbalkens zust�ndig.

'Args: -
'Returns: -
'*************************
Public Sub openStatusBar()

    With maxiBar 'Der Name des Formulars wird programmatisch auf maxiBar gesetzt, daher ist es wichtig, dieses Namensformat zu beachten
    
        .Bar.Width = 0 'Urspr�ngliche Breite des Fortschrittsbalkens - 0
        .Frame.Caption = "0%" 'Die urspr�ngliche Beschreibung des Fortschrittsbalkens
        .Show vbModeless 'Balkenanzeigeparameter - erscheinen nur, wenn sie im Code aufgerufen werden
        
    End With

End Sub

'*************************
'Beschreibung : Diese Methode ist f�r die Anzeigeeinstellungen des Fortschrittsbalkens zust�ndig. Diese Methode wird im Hauptmodul jedes Mal aufgerufen,
'wenn die Anzeige des Fortschrittsbalkens aktualisiert werden muss.

'Args:
'       � step - Integer - Die Sequenznummer des aktuellen Schritts.
'       � totalSteps - Integer - Gesamtzahl der Schritte, die ein Fortschrittsbalken �berwinden muss.
'       � beschreibungString - String - Beschreibung des aktuellen Schritts, die auch auf dem Formular angezeigt wird

'Returns: -
'*************************
Public Sub runStatusBar(step As Integer, totalSteps As Integer, beschreibungString As String)

    With maxiBar
        
        .Bar.Width = 264 * (step / totalSteps)
        .Frame.Caption = Round((step / totalSteps) * 100, 0) & " %"
        .Beschreibung.Caption = beschreibungString
        
    End With

End Sub

'*************************
'Beschreibung : Diese Methode ist daf�r zust�ndig, den Fortschrittsbalken zu beenden, wenn das Programm abgeschlossen ist.

'Args: -
'Returns: -
'*************************
Public Sub deleteBar()

    Unload maxiBar

End Sub




'********************************************************************************************************************
'Anleitung zur Verwendung des MaxiBar-Moduls
'********************************************************************************************************************

'Dieses Modul implementiert, wie die Beschreibung vermuten l�sst, die Funktionalit�t eines vereinfachten und kompakten Fortschrittsbalkens.
'Der einzige Unterschied zwischen diesem und dem vorherigen Fortschrittsbalken besteht darin, dass die Funktionalit�t meines Fortschrittsbalkens in einer zus�tzlichen Klasse implementiert ist und nicht direkt im grafischen Modul.
'Die Klasse ist recht einfach und die Arbeit wird in 3 Methoden durchgef�hrt.
'
'Die folgenden Schritte sollten f�r einen erfolgreichen Einsatz im Projekt befolgt werden:
'
'1) Importieren Sie sowohl das Grafikmodul als auch die Klasse in das Projekt, in dem die Funktionalit�t implementiert werden soll
'2) Erstellen Sie eine Klassenvariable - eine Instanz dieser Klasse im Hauptmodul des Programms
'Private maxiBar As cls_MaxiBar
'3) Definieren Sie im Hauptmodul diese Variable als eine neue Instanz der KlasseSet maxiBar = new cls_maxiBar
'4) F�hren Sie die Methode openStatusBar aus. Mit dieser Methode wird die Startinstanz dieser Klasse erstellt und dem Benutzer angezeigt.maxiBar.openStatusBar
'5) Danach rufen Sie an der Stelle, an der die Fortschrittsanzeige aktualisiert werden soll, die Methode runStatusBar mit den angegebenen Parametern auf - aktuelle Schrittnummer, Gesamtzahl der Schritte,
'   kurze Beschreibung des Schritts.maxiBar.runStatusBar 1, 5, �Step number 1 / 5�
'6) Nach Beendigung der Klasse oder im Falle von errorExit muss die Methode deleteBar aufgerufen werden, um den Fortschrittsbalken zu beenden und das Fenster zu schlie�en. maxiBar.deleteBar
