VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm3 
    Caption         =   "Report Generation Form"
    ClientHeight    =   3040
    ClientLeft      =   110
    ClientTop       =   450
    ClientWidth     =   4580
    OleObjectBlob   =   "UserForm3.frx":0000
    StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' UserForm level dictionary for Pickup Dates by Week
Private dictPickupDatesByWeek As Object

Private Sub UserForm_Initialize()
    Dim dictWeeks As Object
    Dim dictLSP As Object
    Dim wsSource As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim headers As Range, rngData As Range
    Dim colWeek As Long, colLSP As Long, colPickupDate As Long
    Dim i As Long
    Dim weekKey As Variant
    
    Set wsSource = ThisWorkbook.Sheets("LSP_Booking Sheets")
    lastRow = wsSource.Cells(wsSource.Rows.count, 1).End(xlUp).Row
    lastCol = wsSource.Cells(1, wsSource.Columns.count).End(xlToLeft).Column

    Set headers = wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, lastCol))
    Set rngData = wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(lastRow, lastCol))

    ' Find the required columns by header name
    colWeek = Application.Match("Week Number", headers, 0)
    colLSP = Application.Match("LSP (Puninar/KALOG)", headers, 0)
    colPickupDate = Application.Match("Pickup Date", headers, 0)

    If IsError(colWeek) Or IsError(colLSP) Or IsError(colPickupDate) Then
        MsgBox "One or more required columns not found. Please check your headers.", vbCritical
        Exit Sub
    End If

    Set dictWeeks = CreateObject("Scripting.Dictionary")
    Set dictLSP = CreateObject("Scripting.Dictionary")
    Set dictPickupDatesByWeek = CreateObject("Scripting.Dictionary")

    ' Loop through data and build dictionaries
    For i = 1 To rngData.Rows.count
        Dim currentWeek As String
        Dim currentLSP As String
        Dim currentPickupDate As String
        
        currentWeek = CStr(rngData.Cells(i, colWeek).Value)
        currentLSP = CStr(rngData.Cells(i, colLSP).Value)
        
        If IsDate(rngData.Cells(i, colPickupDate).Value) Then
            currentPickupDate = Format(rngData.Cells(i, colPickupDate).Value, "yyyy-mm-dd")
        Else
            currentPickupDate = ""
        End If
        
        ' Add week to dictWeeks
        If currentWeek <> "" Then
            If Not dictWeeks.Exists(currentWeek) Then
                dictWeeks.Add currentWeek, True
            End If
            
            ' Add pickup date under the week key
            If currentPickupDate <> "" Then
                If Not dictPickupDatesByWeek.Exists(currentWeek) Then
                    dictPickupDatesByWeek.Add currentWeek, CreateObject("Scripting.Dictionary")
                End If
                
                If Not dictPickupDatesByWeek(currentWeek).Exists(currentPickupDate) Then
                    dictPickupDatesByWeek(currentWeek).Add currentPickupDate, True
                End If
            End If
        End If
        
        ' Add LSP to dictLSP
        If currentLSP <> "" Then
            If Not dictLSP.Exists(currentLSP) Then
                dictLSP.Add currentLSP, True
            End If
        End If
    Next i

    ' Fill Week listbox
    Me.lstWeek.Clear
    For Each weekKey In dictWeeks.Keys
        Me.lstWeek.AddItem weekKey
    Next weekKey
    
    ' Fill LSP listbox
    Dim lspKey As Variant
    Me.lstLSP.Clear
    For Each lspKey In dictLSP.Keys
        Me.lstLSP.AddItem lspKey
    Next lspKey
    
    ' Fill Pickup Date listbox with ALL unique pickup dates (independent of week selection)
    Dim allPickupDates As Object
    Set allPickupDates = CreateObject("Scripting.Dictionary")
    
    Dim wkKey As Variant
    Dim dtKey As Variant
    
    For Each wkKey In dictPickupDatesByWeek.Keys
        For Each dtKey In dictPickupDatesByWeek(wkKey).Keys
            If Not allPickupDates.Exists(dtKey) Then
                allPickupDates.Add dtKey, True
            End If
        Next dtKey
    Next wkKey
    
    Me.lstPickupDate.Clear
    For Each dtKey In allPickupDates.Keys
        Me.lstPickupDate.AddItem dtKey
    Next dtKey
End Sub

' Remove or comment out lstWeek_Click to prevent filtering Pickup Dates
Private Sub lstWeek_Click()
    ' Intentionally left blank or comment out if you don't want filtering:
    ' UpdatePickupDatesList
End Sub

' Button click to generate report with selected filters
Private Sub btnGenerate_Click()
    Dim selectedWeeks As New Collection
    Dim selectedLSPs As New Collection
    Dim selectedDates As New Collection
    Dim i As Long

    ' Collect selected weeks
    For i = 0 To Me.lstWeek.ListCount - 1
        If Me.lstWeek.Selected(i) Then selectedWeeks.Add Me.lstWeek.List(i)
    Next i

    ' Collect selected LSPs
    For i = 0 To Me.lstLSP.ListCount - 1
        If Me.lstLSP.Selected(i) Then selectedLSPs.Add Me.lstLSP.List(i)
    Next i

    ' Collect selected Pickup Dates
    For i = 0 To Me.lstPickupDate.ListCount - 1
        If Me.lstPickupDate.Selected(i) Then selectedDates.Add Me.lstPickupDate.List(i)
    Next i

    ' Validation
    If selectedLSPs.count = 0 Then
        MsgBox "Please select at least one LSP.", vbExclamation
        Exit Sub
    End If

    If selectedWeeks.count = 0 And selectedDates.count = 0 Then
        MsgBox "Please select at least one Week or one Pickup Date.", vbExclamation
        Exit Sub
    End If

    ' Call your filter/export routine with selected filters
    FilterAndExport_ByWeekLSPAndDate selectedWeeks, selectedLSPs, selectedDates

    MsgBox "Report creation completed.", vbInformation
End Sub
