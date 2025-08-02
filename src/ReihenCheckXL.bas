Attribute VB_Name = "ReihenCheckXL"

Function LückenListe(rng As Range) As String
    Dim i As Long, val1 As Long, val2 As Long
    Dim fehlendeListe As String, zelle As Range
    fehlendeListe = ""

    For Each zelle In rng
        VerwalteTooltipp zelle
        If IsEmpty(zelle) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Leere Zelle: Sie unterbricht die erwartete Zahlenfolge."
            LückenListe = "Sequenzfehler in " & zelle.Address
            Exit Function
        ElseIf Not IsNumeric(zelle.Value) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Typenfehler: Diese Zelle enthält keinen numerischen Wert (z. B. Text)."
            LückenListe = "Typefehler in " & zelle.Address
            Exit Function
        End If
    Next zelle

    For i = 2 To rng.Cells.count
        val1 = CLng(rng.Cells(i - 1).Value)
        val2 = CLng(rng.Cells(i).Value)
        If val2 > val1 + 1 Then
            Dim k As Long
            For k = val1 + 1 To val2 - 1
                fehlendeListe = fehlendeListe & k & ","
            Next k
        End If
    Next i

    If fehlendeListe = "" Then
        fehlendeListe = "Keine Lücken"
    Else
        fehlendeListe = Left(fehlendeListe, Len(fehlendeListe) - 1)
    End If

    LückenListe = fehlendeListe
End Function


Function LückenArray(rng As Range) As Variant
    Dim i As Long, val1 As Long, val2 As Long, lücke As Long
    Dim count As Long, tempArr() As Variant, zelle As Range

    For Each zelle In rng
        VerwalteTooltipp zelle
        If IsEmpty(zelle) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Sequenzfehler: Leere Zelle unterbricht die erwartete Zahlenfolge."
            LückenArray = Array("Sequenzfehler in " & zelle.Address)
            Exit Function
        ElseIf Not IsNumeric(zelle.Value) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Typenfehler: Diese Zelle enthält keinen numerischen Wert."
            LückenArray = Array("Typefehler in " & zelle.Address)
            Exit Function
        End If
    Next zelle

    For i = 2 To rng.Cells.count
        val1 = CLng(rng.Cells(i - 1).Value)
        val2 = CLng(rng.Cells(i).Value)
        If val2 > val1 + 1 Then
            count = count + (val2 - val1 - 1)
        End If
    Next i

    If count = 0 Then
        LückenArray = Array("Keine Lücken")
        Exit Function
    End If

    ReDim tempArr(1 To count)
    count = 1
    For i = 2 To rng.Cells.count
        val1 = CLng(rng.Cells(i - 1).Value)
        val2 = CLng(rng.Cells(i).Value)
        If val2 > val1 + 1 Then
            For lücke = val1 + 1 To val2 - 1
                tempArr(count) = lücke
                count = count + 1
            Next lücke
        End If
    Next i

    LückenArray = Application.Transpose(tempArr)
End Function


Private Sub VerwalteTooltipp(zelle As Range)
    On Error Resume Next
    ' Tooltipp löschen, wenn Zelle korrekt ist
    If Not IsEmpty(zelle) And IsNumeric(zelle.Value) Then
        If Not zelle.Comment Is Nothing Then zelle.ClearComments
    ' Tooltipp setzen bei Fehler
    ElseIf IsEmpty(zelle) Then
        zelle.ClearComments
        zelle.AddComment "Leere Zelle: Sie unterbricht die erwartete Zahlenfolge."
    ElseIf Not IsNumeric(zelle.Value) Then
        zelle.ClearComments
        zelle.AddComment "Typenfehler: Diese Zelle enthält keinen numerischen Wert."
    End If
End Sub





