VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Vertexs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Vertexs
Private mcol As Collection

Public Function Add(No As Long, _
                Optional PT1 As c_Coord, _
                Optional Pt2 As c_Coord, _
                Optional Pt3 As c_Coord) As c_Vertex
    
    Dim objNewMember As c_Vertex
    Set objNewMember = New c_Vertex
    
    objNewMember.No = No
    objNewMember.PT1 = PT1
    objNewMember.Pt2 = Pt2
    objNewMember.Pt3 = Pt3

    mcol.Add objNewMember

    Set Add = objNewMember
    Set objNewMember = Nothing

End Function

Public Property Get item(vntIndexKey As Variant) As c_Vertex
  Set item = mcol(vntIndexKey)
End Property

Public Property Get Count() As Long
    Count = mcol.Count
End Property

Public Property Get Items() As Collection
    Set Items = mcol
End Property

Public Sub Remove(vntIndexKey As Variant)
    mcol.Remove vntIndexKey
End Sub

Private Sub Class_initialize()
    Set mcol = New Collection
End Sub

Private Sub Class_Terminate()
    Set mcol = Nothing
End Sub




