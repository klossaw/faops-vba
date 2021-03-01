Attribute VB_Name = "Ä£¿é1"
Option Explicit
'=======================================================================================
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
  Reverse As String
  Complement As String
  ReverseComplement As String
  RNA As String
  Protein As String
End Type
Public RawSequence() As Sequence
Public RowNumber As Long
Public SeqNumber As Long
Public Code As Object
Public NumberOfCode As Long
'=========================================================================================
' ImportSequence
Public Sub ImportSequence()
'get max rownumbers
RowNumber = ActiveSheet.UsedRange.Rows.Count

'get sequence name and content
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
ReDim RawSequence(1 To SeqNumber) As Sequence
ReDim position(1 To SeqNumber) As Long
For i = 1 To RowNumber
  If regex.test(ActiveSheet.Cells(i, 1)) = True Then
    RawSequence(j + 1).Name = ActiveSheet.Cells(i, 1)
    position(j + 1) = i
    j = j + 1
    End If
    Next i
For i = 1 To SeqNumber - 1
 For j = position(i) + 1 To position(i + 1) - 1
 RawSequence(i).Content = RawSequence(i).Content & ActiveSheet.Cells(j, 1)
 Next j
 Next i
For i = position(SeqNumber) + 1 To RowNumber
 RawSequence(SeqNumber).Content = RawSequence(SeqNumber).Content & ActiveSheet.Cells(i, 1)
 Next i
End Sub
'========================================================================================
Public Sub BasicStatistics()
'Basic statistics of sequence :
'  Length of different sequence;
'  A,C,G,T content of different sequence;
'  GC percentage of diffferent sequence;
Dim OldActive
Set OldActive = ActiveSheet
Call ImportSequence
Dim letter As String, i As Long, j As Long
For i = 1 To SeqNumber
  RawSequence(i).Length = Len(RawSequence(i).Content)
  For j = 1 To RawSequence(i).Length
        letter = Mid(RawSequence(i).Content, j, 1)
        Select Case letter
        Case Is = "A"
            RawSequence(i).CountA = RawSequence(i).CountA + 1
        Case Is = "C"
            RawSequence(i).CountC = RawSequence(i).CountC + 1
        Case Is = "G"
            RawSequence(i).CountG = RawSequence(i).CountG + 1
        Case Is = "T"
            RawSequence(i).CountT = RawSequence(i).CountT + 1
        End Select
        Next j
        RawSequence(i).GCPercentage = (RawSequence(i).CountG + RawSequence(i).CountC) / RawSequence(i).Length
  Next i

'output
Application.DisplayAlerts = False
On Error Resume Next
Worksheets("Basic").Delete
Err.Clear
Application.DisplayAlerts = True
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
Worksheets("Basic").Cells(i + 1, 1) = RawSequence(i).Name
Worksheets("Basic").Cells(i + 1, 2) = RawSequence(i).Content
Worksheets("Basic").Cells(i + 1, 3) = RawSequence(i).Length
Worksheets("Basic").Cells(i + 1, 4) = RawSequence(i).CountA
Worksheets("Basic").Cells(i + 1, 5) = RawSequence(i).CountG
Worksheets("Basic").Cells(i + 1, 6) = RawSequence(i).CountC
Worksheets("Basic").Cells(i + 1, 7) = RawSequence(i).CountT
Worksheets("Basic").Cells(i + 1, 8) = RawSequence(i).GCPercentage
Next i
OldActive.Activate
End Sub

'=============================================================================================
' extract sub_sequence
Public Sub Extract()
Dim OldActive
Set OldActive = ActiveSheet
Dim i As Long
Dim Match As Object
Call ImportSequence
Dim ExtractSequence As String
Dim NextRow As Long, NextCloumn As Long
ExtractSequence = InputBox("Please enter sequence which is to be extracted:  ")
Dim ExtractRegex As Object
Set ExtractRegex = CreateObject("vbscript.regexp")
With ExtractRegex
    .Global = True
    .Pattern = ExtractSequence
    .MultiLine = True
    .ignorecase = True
    End With
Application.DisplayAlerts = False
On Error Resume Next
Worksheets("Fragment").Delete
Err.Clear
Application.DisplayAlerts = True
Worksheets.Add.Name = "Fragment"
Dim OutputSheet As Object
Set OutputSheet = Worksheets("Fragment")
OutputSheet.Cells(1, 1) = "Name of Sequence"
OutputSheet.Cells(1, 2) = "Extracted sequence"
OutputSheet.Cells(1, 3) = "Number of Extracted Sequence "
For i = 1 To SeqNumber
    If ExtractRegex.test(RawSequence(i).Content) = True Then
        Set Match = ExtractRegex.Execute(RawSequence(i).Content)
        NextRow = Cells(Rows.Count, 1).End(xlUp).Row + 1
        OutputSheet.Cells(NextRow, 1) = RawSequence(i).Name
        OutputSheet.Cells(NextRow, 2) = UCase(ExtractSequence)
        OutputSheet.Cells(NextRow, 3) = Match.Count
        End If
    Next i
OldActive.Activate
End Sub

'=================================================================================================
'reverse complement
Public Sub ReverseComplement()
Dim i As Long, j As Long, temp1 As String, temp2 As String
Dim OldActive
Set OldActive = ActiveSheet
Call ImportSequence
Application.DisplayAlerts = False
On Error Resume Next
Worksheets("ReverseComplement").Delete
Err.Clear
Application.DisplayAlerts = True
Worksheets.Add.Name = "ReverseComplement"
Dim OutputSheet
Set OutputSheet = Worksheets("ReverseComplement")
OutputSheet.Cells(1, 1) = " Name of Sequence "
OutputSheet.Cells(1, 2) = " Raw Sequence "
OutputSheet.Cells(1, 3) = " ReverseComplement Sequence"
For i = 1 To SeqNumber
    RawSequence(i).Reverse = StrReverse(RawSequence(i).Content)
    For j = 1 To Len(RawSequence(i).Content)
        temp1 = Mid(RawSequence(i).Content, j, 1)
        temp2 = Mid(RawSequence(i).Reverse, j, 1)
        Select Case temp1
        Case Is = "C"
        temp1 = "G"
        Case Is = "C"
        temp1 = "G"
        Case Is = "T"
        temp1 = "A"
        Case Is = "A"
        temp1 = "T"
        End Select
        Select Case temp2
        Case Is = "C"
        temp2 = "G"
        Case Is = "G"
        temp2 = "C"
        Case Is = "T"
        temp2 = "A"
        Case Is = "A"
        temp2 = "T"
        End Select
        RawSequence(i).Complement = RawSequence(i).Complement & temp1
        RawSequence(i).ReverseComplement = RawSequence(i).ReverseComplement & temp2
        Next j
        OutputSheet.Cells(i + 1, 1) = RawSequence(i).Name
        OutputSheet.Cells(i + 1, 2) = RawSequence(i).Content
        OutputSheet.Cells(i + 1, 3) = RawSequence(i).ReverseComplement
        Next i
OldActive.Activate
End Sub
'===================================================================================================
'Find motif
Public Sub FindMotif()
Call ImportSequence
Dim OldActive
Set OldActive = ActiveSheet
Application.DisplayAlerts = False
On Error Resume Next
Worksheets("Motif").Delete
Err.Clear
Application.DisplayAlerts = True
Worksheets.Add.Name = "Motif"
Dim OutputSheet
Set OutputSheet = Worksheets("Motif")
OutputSheet.Cells(1, 1) = "Name of Sequence"
OutputSheet.Cells(1, 2) = "Motif"
OutputSheet.Cells(1, 3) = "Location"
Dim Motif As String, output
Motif = InputBox("Please enter a motif:  ")
Dim MotifRegex, Match
Set MotifRegex = CreateObject("vbscript.regexp")
With MotifRegex
    .Global = True
    .Pattern = Motif
    .ignorecase = True
    End With
Dim i As Long, j As Long
For i = 1 To SeqNumber
    If MotifRegex.test(RawSequence(i).Content) = True Then
        Set Match = MotifRegex.Execute(RawSequence(i).Content)
        For j = 1 To Match.Count
        output = output & "   " & Match(j - 1).firstindex + 1
        Next j
        OutputSheet.Cells(i + 1, 1) = RawSequence(i).Name
        OutputSheet.Cells(i + 1, 2) = UCase(Motif)
        OutputSheet.Cells(i + 1, 3) = output
        output = ""
        End If
        Next i
OldActive.Activate
End Sub
'=======================================================================================================
' DNA-Into-RNA
Public Sub DnaToRna()
Call ImportSequence
Dim OldActive
Set OldActive = ActiveSheet
Application.DisplayAlerts = False
On Error Resume Next
Worksheets("RNA").Delete
Worksheets.Add.Name = "RNA"
Err.Clear
Application.DisplayAlerts = True

Dim OutputSheet
Set OutputSheet = Worksheets("RNA")
OutputSheet.Cells(1, 1) = "Name of Sequence"
OutputSheet.Cells(1, 2) = "RNA"
Dim temp As String, Match
Dim i As Long, j As Long
Dim RnaRegex As Object
Set RnaRegex = CreateObject("vbscript.regexp")
With RnaRegex
    .Global = True
    .Pattern = "T"
    .ignorecase = True
    End With
For i = 1 To SeqNumber
        RawSequence(i).RNA = RnaRegex.Replace(RawSequence(i).Content, "U")
        OutputSheet.Cells(i + 1, 1) = RawSequence(i).Name
        OutputSheet.Cells(i + 1, 2) = RawSequence(i).RNA
        Next i
OldActive.Activate
End Sub
'=========================================================================================================
'RNA-into-protein
Public Sub RnaToProtein()
Call ImportSequence
Call DnaToRna
Dim OldActive
Set OldActive = ActiveSheet
On Error Resume Next
Worksheets("Protein").Delete
Worksheets.Add.Name = "Protein"
Err.Clear
Set Code = CreateObject("scripting.dictionary")
    Code.Add "GCU", "A"
    Code.Add "GCC", "A"
    Code.Add "GCA", "A"
    Code.Add "GCG", "A"
    Code.Add "CGU", "R"
    Code.Add "CGC", "R"
    Code.Add "CGA", "R"
    Code.Add "CGG", "R"
    Code.Add "AGA", "R"
    Code.Add "AGG", "R"
    Code.Add "AAU", "N"
    Code.Add "AAC", "N"
    Code.Add "GAU", "D"
    Code.Add "GAC", "D"
    Code.Add "UGU", "C"
    Code.Add "UGC", "C"
    Code.Add "CAA", "Q"
    Code.Add "CAG", "Q"
    Code.Add "CAU", "H"
    Code.Add "CAC", "H"
    Code.Add "AUG", "M"
    Code.Add "AUU", "I"
    Code.Add "AUC", "I"
    Code.Add "AUA", "I"
    Code.Add "CUU", "L"
    Code.Add "CUC", "L"
    Code.Add "CUA", "L"
    Code.Add "CUG", "L"
    Code.Add "UUA", "L"
    Code.Add "UUG", "L"
    Code.Add "AAA", "K"
    Code.Add "AAG", "K"
    Code.Add "UUU", "F"
    Code.Add "UUC", "F"
    Code.Add "CCU", "P"
    Code.Add "CCC", "P"
    Code.Add "CCA", "P"
    Code.Add "CCG", "P"
    Code.Add "UCU", "S"
    Code.Add "UCC", "S"
    Code.Add "UCA", "S"
    Code.Add "UCG", "S"
    Code.Add "AGU", "S"
    Code.Add "AGC", "S"
    Code.Add "ACU", "T"
    Code.Add "ACC", "T"
    Code.Add "ACA", "T"
    Code.Add "ACG", "T"
    Code.Add "UGG", "T"
    Code.Add "UAU", "Y"
    Code.Add "UAC", "Y"
    Code.Add "GUU", "V"
    Code.Add "GUC", "V"
    Code.Add "GUA", "V"
    Code.Add "GUG", "V"
    Code.Add "UAA", "STOP"
    Code.Add "UGA", "STOP"
    Code.Add "UAG", "STOP"
    Code.Add "TGG", "W"
    Code.Add "GAA", "E"
    Code.Add "GAG", "E"
    Code.Add "GGT", "G"
    Code.Add "GGC", "G"
    Code.Add "GGA", "G"
    Code.Add "GGG", "G"
NumberOfCode = Code.Count
Dim temp As String
Dim i As Long
Dim j As Long
Dim OutputSheet
Set OutputSheet = Worksheets("Protein")
OutputSheet.Cells(1, 1) = "Name of sequence"
OutputSheet.Cells(1, 2) = "Protein sequence"
For i = 1 To SeqNumber
    For j = 1 To Len(RawSequence(i).RNA) Step 3
    temp = Mid(RawSequence(i).RNA, j, 3)
    If Code(temp) <> "STOP" Then
    RawSequence(i).Protein = RawSequence(i).Protein & Code(temp)
    End If
    Next j
    OutputSheet.Cells(i + 1, 1) = RawSequence(i).Name
    OutputSheet.Cells(i + 1, 2) = RawSequence(i).Protein
    Next i
    OldActive.Activate
End Sub
'==================================================================================================
'inferring mRNA from protein

