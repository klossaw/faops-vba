Attribute VB_Name = "Ä£¿é1"
Option Explicit
'define a new data type:Sequence
Type Sequence
  Name As String
  Content As String
  Length As Long
  CountA As Long
  CountC As Long
  CountG As Long
  CountT As Long
  GCPercentage As Double
End Type
Public RowSequence() As Sequence
Public RowNumber As Long
Public Sub ImportSequence()
'get max rownumbers
RowNumber = ActiveSheet.UsedRange.Rows.Count

'get sequence name and content
Dim SeqNumber As Long
Dim position() As Long
Dim regex As Object
SeqNumber = 0
Set regex = CreateObject("vbscript.regexp")
With regex
  .Global = True
  .Pattern = ">"
End With
Dim i As Long
Dim j As Long
j = 0
For i = 1 To RowNumber
  If regex.test(ActiveSheet.Cells(i, 1)) = True Then
    SeqNumber = SeqNumber + 1
    End If
  Next i
ReDim RowSequence(1 To SeqNumber) As Sequence
ReDim position(1 To SeqNumber) As Long
For i = 1 To RowNumber
  If regex.test(ActiveSheet.Cells(i, 1)) = True Then
    RowSequence(j + 1).Name = ActiveSheet.Cells(i, 1)
    position(j + 1) = i
    j = j + 1
    End If
    Next i
For i = 1 To SeqNumber - 1
 For j = position(i) + 1 To position(i + 1) - 1
 RowSequence(i).Content = RowSequence(i).Content & ActiveSheet.Cells(j, 1)
 Next j
 Next i
For i = position(SeqNumber) + 1 To RowNumber
 RowSequence(SeqNumber).Content = RowSequence(SeqNumber).Content & ActiveSheet.Cells(i, 1)
 Next i

'basic statistics of sequence
Dim letter As String
For i = 1 To SeqNumber
  RowSequence(i).Length = Len(RowSequence(i).Content)
  For j = 1 To RowSequence(i).Length
        letter = Mid(RowSequence(i).Content, j, 1)
        Select Case letter
        Case Is = "A"
            RowSequence(i).CountA = RowSequence(i).CountA + 1
        Case Is = "C"
            RowSequence(i).CountC = RowSequence(i).CountC + 1
        Case Is = "G"
            RowSequence(i).CountG = RowSequence(i).CountG + 1
        Case Is = "T"
            RowSequence(i).CountT = RowSequence(i).CountT + 1
        End Select
        Next j
        RowSequence(i).GCPercentage = (RowSequence(i).CountG + RowSequence(i).CountC) / RowSequence(i).Length
  Next i

'output
Worksheets.Add.Name = "Basic"
Worksheets("Basic").Cells(1, 1) = "Sequence name"
Worksheets("Basic").Cells(1, 2) = "Sequence"
Worksheets("Basic").Cells(1, 3) = "Length of Sequence"
Worksheets("Basic").Cells(1, 4) = "Number of Base A"
Worksheets("Basic").Cells(1, 5) = "Number of Base G"
Worksheets("Basic").Cells(1, 6) = "Number of Base C"
Worksheets("Basic").Cells(1, 7) = "Number of Base T"
Worksheets("Basic").Cells(1, 8) = "GC percentage"
For i = 1 To SeqNumber
Worksheets("Basic").Cells(i + 1, 1) = RowSequence(i).Name
Worksheets("Basic").Cells(i + 1, 2) = RowSequence(i).Content
Worksheets("Basic").Cells(i + 1, 3) = RowSequence(i).Length
Worksheets("Basic").Cells(i + 1, 4) = RowSequence(i).CountA
Worksheets("Basic").Cells(i + 1, 5) = RowSequence(i).CountG
Worksheets("Basic").Cells(i + 1, 6) = RowSequence(i).CountC
Worksheets("Basic").Cells(i + 1, 7) = RowSequence(i).CountT
Worksheets("Basic").Cells(i + 1, 8) = RowSequence(i).GCPercentage
Next i
End Sub
