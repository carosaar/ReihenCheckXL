Attribute VB_Name = "ReihenCheckXL"

Function L�ckenListe(rng As Range) As String
    Dim i As Long, val1 As Long, val2 As Long
    Dim fehlendeListe As String, zelle As Range
    fehlendeListe = ""

    For Each zelle In rng
        VerwalteTooltipp zelle
        If IsEmpty(zelle) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Leere Zelle: Sie unterbricht die erwartete Zahlenfolge."
            L�ckenListe = "Sequenzfehler in " & zelle.Address
            Exit Function
        ElseIf Not IsNumeric(zelle.Value) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Typenfehler: Diese Zelle enth�lt keinen numerischen Wert (z.�B. Text)."
            L�ckenListe = "Typefehler in " & zelle.Address
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
        fehlendeListe = "Keine L�cken"
    Else
        fehlendeListe = Left(fehlendeListe, Len(fehlendeListe) - 1)
    End If

    L�ckenListe = fehlendeListe
End Function


Function L�ckenArray(rng As Range) As Variant
    Dim i As Long, val1 As Long, val2 As Long, l�cke As Long
    Dim count As Long, tempArr() As Variant, zelle As Range

    For Each zelle In rng
        VerwalteTooltipp zelle
        If IsEmpty(zelle) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Sequenzfehler: Leere Zelle unterbricht die erwartete Zahlenfolge."
            L�ckenArray = Array("Sequenzfehler in " & zelle.Address)
            Exit Function
        ElseIf Not IsNumeric(zelle.Value) Then
            On Error Resume Next: zelle.ClearComments: zelle.AddComment "Typenfehler: Diese Zelle enth�lt keinen numerischen Wert."
            L�ckenArray = Array("Typefehler in " & zelle.Address)
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
        L�ckenArray = Array("Keine L�cken")
        Exit Function
    End If

    ReDim tempArr(1 To count)
    count = 1
    For i = 2 To rng.Cells.count
        val1 = CLng(rng.Cells(i - 1).Value)
        val2 = CLng(rng.Cells(i).Value)
        If val2 > val1 + 1 Then
            For l�cke = val1 + 1 To val2 - 1
                tempArr(count) = l�cke
                count = count + 1
            Next l�cke
        End If
    Next i

    L�ckenArray = Application.Transpose(tempArr)
End Function


Private Sub VerwalteTooltipp(zelle As Range)
    On Error Resume Next
    ' Tooltipp l�schen, wenn Zelle korrekt ist
    If Not IsEmpty(zelle) And IsNumeric(zelle.Value) Then
        If Not zelle.Comment Is Nothing Then zelle.ClearComments
    ' Tooltipp setzen bei Fehler
    ElseIf IsEmpty(zelle) Then
        zelle.ClearComments
        zelle.AddComment "Leere Zelle: Sie unterbricht die erwartete Zahlenfolge."
    ElseIf Not IsNumeric(zelle.Value) Then
        zelle.ClearComments
        zelle.AddComment "Typenfehler: Diese Zelle enth�lt keinen numerischen Wert."
    End If
End Sub





