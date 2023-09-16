Attribute VB_Name = "Transform"
Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = "M\n14"
    'Initialisation des variables
    
    Dim File_Name As String, File_Final As String, Letter_Name1 As String, Letter_Search1 As String, Letter_Name2 As String, Letter_Search2 As String
    Dim SearchName1 As String, SearchName2 As String
    Dim Position_max As Long, Value_Id_Ligne As Long
    
    File_Name = "Transform"
    File_Final = "Result"
    Position_max = 54629
    
    Letter_Name1 = "H"
    Letter_Search1 = "A"
    Letter_Name2 = "I"
    Letter_Search2 = "B"
    
    Value_Id_Ligne = 1
    
    SearchName1 = "Total"
    SearchName2 = "Section"
    
    
    
    'Appel de la fonction
    
    SectionSearch File_Name, SearchName2, Position_max
    Compare File_Name, SearchName1, Position_max
    
    Do While Position_max >= Value_Id_Ligne
        Value_Id_Ligne = Search(File_Name, Letter_Name1, Letter_Name2, Letter_Search1, Letter_Search2, Value_Id_Ligne)
    Loop
    
    Duplique File_Name, File_Final
    Sheets(File_Final).Select
    Range("A2:A65000").SpecialCells(xlCellTypeBlanks).EntireRow.Delete
    
End Sub





Private Sub SectionSearch(ByVal File_Name As String, ByVal SearchName As String, ByVal Position_max As Long)

    Sheets(File_Name).Select
    ma_case1 = "A"
    Position = 1
    Name_Cellule = ""
    
   Do While Position <= Position_max
    
        ma_case = ma_case1 & Position
        ma_chaine = Sheets(File_Name).Range(ma_case).Value
        Pos = InStr(1, ma_chaine, SearchName)
        
        If Pos = 1 Then
            
            Cellule = "B" & Position
            Name_Cellule = Sheets(File_Name).Range(Cellule).Value
            Rows(Position).Select
            Selection.ClearContents
            
        End If
        Position = Position + 1
        Sheets(File_Name).Range("J" & Position) = Name_Cellule
   Loop
    
End Sub



Private Function Search(ByVal File_Name As String, ByVal Letter_Name1 As String, ByVal Letter_Name2 As String, ByVal Letter_Search1 As String, ByVal Letter_Search2 As String, Lig As Long) As Long
Attribute Search.VB_ProcData.VB_Invoke_Func = "m\n14"
 
    
    Do While Not IsEmpty(Range("A" & Lig))
        Lig = Lig + 1
    Loop

    Id_Value_1 = Sheets(File_Name).Range(Letter_Search1 & (Lig + 1)).Value
    Id_Value_2 = Sheets(File_Name).Range(Letter_Search2 & (Lig + 1)).Value
    Id_Ligne_Start = (Lig + 2)
    Lig = Id_Ligne_Start
    
    Rows(Id_Ligne_Start - 1).Select
    Selection.ClearContents
    
    Do While Not IsEmpty(Range("A" & Lig))
        Lig = Lig + 1
    Loop
    
    Id_Ligne_End = (Lig - 1)
    Sheets(File_Name).Range(Letter_Name1 & Id_Ligne_Start & ":" & Letter_Name1 & Id_Ligne_End) = Id_Value_1
    Sheets(File_Name).Range(Letter_Name2 & Id_Ligne_Start & ":" & Letter_Name2 & Id_Ligne_End) = Id_Value_2
    Search = Lig - 1
    
End Function



Private Sub Duplique(ByVal File_Copy As String, ByVal File_Paste As String)
    
    Sheets.Add(After:=Worksheets(Worksheets.Count)).Name = File_Paste
    
    Sheets(File_Copy).Select
    Columns("J:J").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("A:A").Select
    ActiveSheet.Paste
    
    Sheets(File_Copy).Select
    Columns("H:H").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("B:B").Select
    ActiveSheet.Paste
    
    Sheets(File_Copy).Select
    Columns("I:I").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("C:C").Select
    ActiveSheet.Paste
    
    Sheets(File_Copy).Select
    Columns("A:A").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("D:D").Select
    ActiveSheet.Paste
    
    
    Sheets(File_Copy).Select
    Columns("B:B").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("E:E").Select
    ActiveSheet.Paste
    
    
    Sheets(File_Copy).Select
    Columns("C:C").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("F:F").Select
    ActiveSheet.Paste
    
    
    Sheets(File_Copy).Select
    Columns("D:D").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("G:G").Select
    ActiveSheet.Paste
    
    
    Sheets(File_Copy).Select
    Columns("E:E").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("H:H").Select
    ActiveSheet.Paste
    
    
    Sheets(File_Copy).Select
    Columns("F:F").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("I:I").Select
    ActiveSheet.Paste
    
    
    Sheets(File_Copy).Select
    Columns("G:G").Select
    Selection.Copy
    Sheets(File_Paste).Select
    Columns("J:J").Select
    ActiveSheet.Paste
    
End Sub



Private Sub Compare(ByVal File_Name As String, ByVal SearchName As String, ByVal Position_max As Long)

    Sheets(File_Name).Select
    ma_case1 = "A"
    Position = 1
    
   Do While Position <= Position_max
    
        ma_case = ma_case1 & Position
        ma_chaine = Sheets(File_Name).Range(ma_case).Value
        Pos = InStr(1, ma_chaine, SearchName)
        
        If Pos = 1 Then
            
            ' Cellule = Position & ":" & (Position + 1)
            Cellule = Position
            Rows(Cellule).Select
            Selection.ClearContents
            
        End If
        Position = Position + 1
   Loop
    
    
    Cellule = Position & ":" & (Position + 100)
    Rows(Cellule).Select
    Selection.ClearContents
    
End Sub