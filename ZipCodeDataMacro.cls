VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim Zip_Code_Map, Neighborhood_Map As Scripting.Dictionary


Sub run()

  
    Dim Visitors As Scripting.Dictionary
    Set Visitors = New Scripting.Dictionary
    Dim i, j As Integer
    i = 1
    j = 1
    Do While True
        Dim value As String
        value = Cells(i, j).value
        If value = "" Then
            Exit Do
        Else:
            Dim Name As String
            Dim vis As Integer
            Name = Binary_Town_Name(value)
            neighbor = Binary_Neighborhood(Name)
            vis = Cells(i, j + 1)
            'MsgBox (name & " (" & Neighborhood(name) & ")")
            If Visitors.Exists(neighbor) Then
                Visitors.Item(neighbor) = Visitors.Item(neighbor) + vis
            Else:
                Visitors.Add neighbor, vis
            End If
            
        End If
        i = i + 1
    Loop
    
    Sheets("Output").Select
    Cells.Clear
    i = 1
    j = 1
    For Each strKey In Visitors.Keys()
        Cells(i, j) = strKey
        Cells(i, j + 1) = Visitors.Item(strKey)
        i = i + 1
    Next
    
End Sub

Function readZips() As Scripting.Dictionary
    If (Zip_Code_Map.Count > 0) Then
        readZips = Zip_Code_Map
    End If
    Sheets("Zip_Code_Directory").Select
    Dim zip_map As Scripting.Dictionary
    Set zip_map = New Scripting.Dictionary
    
    Dim i, j As Integer
    i = 1
    j = 1
    Do While True
        Dim value As String
        value = Cells(i, j).value
        If value = "" Then
            Exit Do
        Else:
            zip_map.Add value, Cells(i, j + 1)
        End If
        i = i + 1
    Loop
    Set Zip_Code_Map = zip_map
    readZips = zip_map
    
    Sheets("Input").Select
End Function

Function readNeighborhoods() As Scripting.Dictionary
    If (Neighborhood_Map.Count > 0) Then
        readNeighborhoods = Neighborhood_Map
    End If
    Sheets("Neighborhoods").Select
    Dim neighs As Scripting.Dictionary
    Set neighs = New Scripting.Dictionary
    
    Dim i, j As Integer
    i = 1
    j = 1
    Do While True
        Dim value As String
        value = Cells(i, j).value
        If value = "" Then
            Exit Do
        Else:
            neighs.Add value, Cells(i, j + 1)
        End If
        i = i + 1
    Loop
    Set Neighborhood_Map = neighs
    readNeighborhoods = neighs
    
    Sheets("Input").Select
End Function

Function Binary_Town_Name(Zip_Code As String)
    Dim zip As Scripting.Dictionary
    Set zip = readZips()
    If zip.Exists(Zip_Code) Then
        Binary_Town_Name = zip.Item(Zip_Code)
    Else:
        If (Len(Zip_Code) = 5) And IsNumeric(Zip_Code) Then
                Binary_Town_Name = "Out of State"
        Else:
                Binary_Town_Name = "Unknown"
        End If
    End If
    
End Function

Function Binary_Neighborhood(Name As String)
    Dim neigh As Scripting.Dictionary
    Set neigh = readNeighborhoods()
    If neigh.Exists(Name) Then
        Binary_Neighborhood = neigh.Item(Name)
    Else:
        If Town_Name = "Unknown" Or Town_Name = "Out of State" Then
            Binary_Neighborhood = Town_Name
        Else:
            Binary_Neighborhood = "Misc."
        End If
    End If
    
End Function

Function Town_Name(Zip_Code As String, Zip_Code_Map As Scripting.Dictionary)
    Sheets("Zip_Code_Directory").Select
    If (IsNumeric(Zip_Code)) Then
        Dim i, j As Integer
        i = 1
        j = 1
        Do While True
            Dim value As String
            value = Cells(i, j).value
            If value = "" Then
                If (Len(Zip_Code) = 5) Then
                    Town_Name = "Out of State"
                Else:
                    Town_Name = "Unknown"
                End If
                Exit Do
            ElseIf (value = Zip_Code) Then
                Town_Name = Cells(i, j + 1)
                Exit Do
            End If
            i = i + 1
        Loop
    Else:
        Town_Name = "Unknown"
    End If
    
    Sheets("Input").Select
End Function

Function Neighborhood(Town_Name As String)
    Sheets("Neighborhoods").Select
    Dim i, j As Integer
    i = 1
    j = 1
    Do While True
        Dim value As String
        value = Cells(i, j).value
        If value = "" Then
            If Town_Name = "Unknown" Or Town_Name = "Out of State" Then
                Neighborhood = Town_Name
                Exit Do
            End If
            Neighborhood = "Misc."
            Exit Do
        ElseIf (value = Town_Name) Then
            Neighborhood = Cells(i, j + 1)
            Exit Do
        End If
        i = i + 1
    Loop
    
    Sheets("Input").Select
End Function

