VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "c_Vertex"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'c_Vertex

Private mvarNo As Long
Private mvarPt1 As c_Coord
Private mvarPt2 As c_Coord
Private mvarPt3 As c_Coord

Public Property Let No(ByVal vData As Long)
    mvarNo = vData
End Property

Public Property Get No() As Long
    No = mvarNo
End Property

Public Property Let PT1(vData As c_Coord)
    Set mvarPt1 = vData
End Property

Public Property Get PT1() As c_Coord
    Set PT1 = mvarPt1
End Property

Public Property Let Pt2(vData As c_Coord)
    Set mvarPt2 = vData
End Property

Public Property Get Pt2() As c_Coord
    Set Pt2 = mvarPt2
End Property

Public Property Let Pt3(vData As c_Coord)
    Set mvarPt3 = vData
End Property

Public Property Get Pt3() As c_Coord
    Set Pt3 = mvarPt3
End Property


