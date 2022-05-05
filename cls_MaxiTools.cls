VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cls_MaxiTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=====================================================================
' Version: 1.0
' Datum: 17.08.2021
' Autor: M. Shtekhno
'--------------------------------------------------
' Beschreibung:
'--------------------------------------------------
' • Klasse mit Funktionen und Werkzeugen, die immer wieder Verwendung finden
' • Die Nachvollziehbarkeit des Codes bleibt bei mir im Priorität. Hier werden also nur die Methoden gespeichert, die leicht zu verstehen und intuitiv klar sind.
'--------------------------------------------------
' Update Log: 1.* -> 1.*
'--------------------------------------------------
' • Beschrebung :
'====================================================================================================================
Option Explicit

'********************************************************************************************************************
'Klassenvariablen
'********************************************************************************************************************
Private cls_arr_AppSettings As Variant

'Mit den folgenden 3 Funktionen kann man die Ausführung des Makros deutlich beschleunigen.

'*************************
'Beschreibung : Speichert und schaltet aktuelle Applikationseinstellungen aus
'Input: -
'Output: -
'*************************
Public Sub disableAppSettings()

    Call f_ChangeAppSettings(True)
    
End Sub

'*************************
'Beschreibung : Schaltet zuvor gespeicherte Applikationseinstellungen wieder ein
'Input: -
'Output: -
'*************************
Public Sub enableAppSettings()
    
    Call f_ChangeAppSettings(False)
    
End Sub

'*************************
'Beschreibung : Diese Methode löscht das VBA Modul / Klassenmodul aus dem Projekt
'Input: Workbook aus welcher das Modul gelöscht werden muss, Der Name des Moduls
'Output: -
'*************************
Public Sub removeComponent(Workbook As Workbook, Name As String)
    
    Workbook.VBProject.VBComponents.Remove Workbook.VBProject.VBComponents(Name)
    
End Sub
'*************************
'Beschreibung : Diese Methode importiert das VBA Modul / Klassenmodul in das Projekt
'Input: Workbook in welcher das Modul importiert werden muss, Path zu den Ordner wo das Modul liegt, Der Name des Moduls
'Output: -
'*************************
Public Sub addComponent(Workbook As Workbook, Path As String, Name As String)

    With Workbook.VBProject
    .VBComponents.Import Path & "\" & Name
    End With
    
End Sub
'*************************
'Beschreibung : Diese Methode blendet sogar das VeryHidden Sheet ein
'Input: Name des Sheets, das eingeblenden werden muss
'Output: -
'*************************
Public Sub visible(Sheetname As String)

    If Sheets(Sheetname).visible = xlSheetVeryHidden Then
        Sheets(Sheetname).visible = xlSheetVisible
    Else
        Sheets(Sheetname).visible = xlSheetVisible
    End If
    
End Sub
'*************************
'Beschreibung : Diese Funktion setzt das Code in Modul (VBA) ein
'Input: Name des Reiters, in weicher soll das VBA Code eingesetzt werden, Codezeile as String, Workbook
'Output: -
'*************************
Public Sub SetSheetsVBACode(Worksheet As String, CodeString As String, Workbook As Workbook)
With Workbook.VBProject.VBComponents(Worksheets(Worksheet).CodeName).CodeModule
    If .CountOfLines > 0 Then
        .DeleteLines 1, .CountOfLines
    End If
    If CodeString <> "" Then
        .CodePane.CodeModule.AddFromString CodeString
    End If
End With
End Sub
'*************************
'Beschreibung : Speichert aktuelle Einstellungen ab und schaltet sie aus oder setzt gespeicherte Einstellungen wieder ein (je nach bool_DisableAppSettings)
'Input: Boolische Variable: True für Enable, False für Disable
'Anmerkung: wird im Zusammenhang mit den disableAppSettings, enableAppSettings Methoden benutzt
'*************************
Private Sub f_ChangeAppSettings( _
    ByVal bool_DisableAppSettings As Boolean)
    
    If bool_DisableAppSettings Then
    
        cls_arr_AppSettings = Array( _
                              Application.ScreenUpdating, _
                              Application.Calculation, _
                              Application.EnableEvents, _
                              Application.StatusBar, _
                              Application.DisplayAlerts)
        
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
        Application.EnableEvents = False
        Application.StatusBar = False
        Application.DisplayAlerts = False
    
    Else
    
        On Error Resume Next
        
        Application.ScreenUpdating = cls_arr_AppSettings(0)
        Application.Calculation = cls_arr_AppSettings(1)
        Application.EnableEvents = cls_arr_AppSettings(2)
        Application.StatusBar = cls_arr_AppSettings(3)
        Application.DisplayAlerts = cls_arr_AppSettings(4)
        
        If Err.Number <> 0 Then GoTo errorOccured
        
        On Error GoTo 0
        
    End If
    
    Exit Sub

errorOccured:
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.StatusBar = True
    Application.DisplayAlerts = True
    
    Err.Raise 1
    
End Sub

'********************************************************************************************************************
'Methoden
'********************************************************************************************************************

'*************************
'Beschreibung : Diese Funktion prüfft ob ein Blatt unter dem Reiternamen Sheetname in Arbeitsmappe wb zu finden ist
'Input: Name des Sheets, Workbook zu untersuchen
'Output: True wenn das Sheet existiert, False wenn nicht
'*************************
Public Function sheetExists( _
    ByVal Sheetname As String, _
    ByRef wb As Workbook) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    
    ' Sollte das Blatt nicht existieren, wird ein Fehler provoziert
    '-----------------------------------------
    Set ws = wb.Worksheets(Sheetname)
    
    If Err.Number = 0 Then
        sheetExists = True
    End If
    
    On Error GoTo 0
    
End Function
'*************************
'Beschreibung : Diese Funktion returns das Index der letzten Zeile in gegebener Spalte
'Input: Worksheet, Name der Spalte (String, also z.B. "A", "G", "AX" usw.)
'Output: Index der letzten Zeile as Integer
'*************************
Public Function getLastRow( _
    ByRef Worksheet As Worksheet, _
    ByVal Column As String _
    ) As Integer

    getLastRow = Worksheet.Cells(Worksheet.Rows.Count, Column).End(xlUp).Row
    
End Function
'*************************
'Beschreibung : Diese Funktion returns der Name der letzten Spalte in gegebener Zeile
'Input: Worksheet, Index der Zeile
'Output: Name der letzten Spalte (String, also z.B. "A", "G", "AX" usw.)
'*************************
Public Function getLastColumn( _
    ByRef Worksheet As Worksheet, _
    ByVal Row As Long _
    ) As String

    getLastColumn = Split(Cells(1, Worksheet.Cells(Row, Worksheet.Columns.Count).End(xlToLeft).Column).Address, "$")(1)

End Function
'*************************
'Beschreibung : Diese Funktion returns der Name der letzten Spalte in gegebener Zeile
'Input: Worksheet, Index der Zeile
'Output: Inedx der letzten Spalte (Integer, also z.B. 10)
'*************************
Public Function getLastColumnAsNumber( _
    ByRef Worksheet As Worksheet, _
    ByVal Row As Long _
    ) As Integer

    getLastColumnAsNumber = Worksheet.Cells(Row, Worksheet.Columns.Count).End(xlToLeft).Column

End Function
'*************************
'Beschreibung : Öffnet das Auswahlfenster für Dateien
'Input: Boolische Variable: True - Auswahl von mehreren Dateien möglich, False - nicht Möglich, String - Titel für ein MsgBox
'Output: Gibt ein Array mit ausgewählten Dateipfaden zurück. Sollte nichts ausgewählt worden sein, ist UBound(f_Userselection_File) = -1
'*************************
Public Function selectFiles( _
       Optional ByVal bool_Multifile As Boolean, _
       Optional ByVal str_Name As String = "Datei")
       
    ' Initialisiere arr_tmp als leeres Array
    Dim arr_tmp As Variant: arr_tmp = Array()
    Dim item_Path As Variant
    
    ' Öffne Auswahlfenster
    '-----------------------------------------
    With Application.FileDialog(msoFileDialogFilePicker)
        .Title = "Bitte " & str_Name & " wählen"
        .AllowMultiSelect = bool_Multifile
        .InitialFileName = ThisWorkbook.Path
        .Show
        For Each item_Path In .SelectedItems
            ReDim Preserve arr_tmp(UBound(arr_tmp) + 1)
            arr_tmp(UBound(arr_tmp)) = item_Path
        Next item_Path
    End With
    
    selectFiles = arr_tmp
    
End Function
'*************************
'Beschreibung : Diese Funktion kopiert das Code des Moduls (VBA)
'Input: Names des Reiters, aus welcher das VBA Code kopiert werden muss
'Output: VBA Code des Sheets as a String
'*************************
Public Function GetSheetsVBACode(Worksheet As String)
With ThisWorkbook.VBProject.VBComponents(Worksheets(Worksheet).CodeName).CodeModule
    If .CountOfLines > 0 Then
        GetSheetsVBACode = .Lines(1, .CountOfLines)
    End If
End With
End Function











