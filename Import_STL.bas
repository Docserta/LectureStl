Attribute VB_Name = "Import_STL"


Option Explicit

Sub CATMain()
'Trace les triangles défini dans un fichier STL
Dim oVertexs As c_Vertexs
Dim isBinaire  As Boolean
'Initialisation des classes
    Set oVertexs = New c_Vertexs

'Ouvre la boite de dialogue
    Load Frm_Demarrage
    Frm_Demarrage.Show
    If Not Frm_Demarrage.ChB_OkAnnule Then
        End
    End If
    FicSTL = Frm_Demarrage.TB_Fichier
    If FicSTL = "" Then End 'on vérifie que qque chose a bien été selectionné
    If Frm_Demarrage.Rbt_BIN = True Then
        isBinaire = True
    Else
        isBinaire = False
    End If
    ValSeuil = CDbl(Frm_Demarrage.CBL_Seuil)
    DecoupFic = Frm_Demarrage.ChB_Decoup
    NbItemDecoup = CLng(Frm_Demarrage.CBL_NbPt)
    Unload Frm_Demarrage
    
'Collecte les vertex
    If isBinaire Then
        Set oVertexs = lectureSTLBinaire(FicSTL)
    Else
        Set oVertexs = ColSTL(FicSTL)
    End If
'Tracé des mailles
    CreateMailles oVertexs

'Libération des objets
    Set oVertexs = Nothing

End Sub

Private Function ColSTL(FicSTL) As c_Vertexs
'Collecte les vertex dans le fichier STL
Dim oVertex As c_Vertex
Dim oVertexs As c_Vertexs
Dim Pt As c_Coord
Dim f, fs
Dim CurLig As String
Dim cpt As Long

'Initialisation des classes
    Set oVertex = New c_Vertex
    Set oVertexs = New c_Vertexs
    Set fs = CreateObject("scripting.filesystemobject")

    Set f = fs.opentextfile(FicSTL, ForReading, 1)
    Do While Not f.AtEndOfStream
        cpt = cpt + 1
        CurLig = f.ReadLine
        If InStr(1, CurLig, "outer loop", vbTextCompare) > 0 Then
            CurLig = f.ReadLine
            If InStr(1, CurLig, "vertex", vbTextCompare) > 0 Then
                Set Pt = AdPt(CurLig)
                oVertex.Pt1 = Pt
            End If
            CurLig = f.ReadLine
            If InStr(1, CurLig, "vertex", vbTextCompare) > 0 Then
                Set Pt = AdPt(CurLig)
                oVertex.Pt2 = Pt
            End If
            CurLig = f.ReadLine
            If InStr(1, CurLig, "vertex", vbTextCompare) > 0 Then
                Set Pt = AdPt(CurLig)
                oVertex.Pt3 = Pt
            End If
            oVertexs.Add cpt, oVertex.Pt1, oVertex.Pt2, oVertex.Pt3
        End If
    Loop
Set ColSTL = oVertexs
'Liberation des classes
    Set oVertex = Nothing
    Set oVertexs = Nothing
    Set f = Nothing
    Set fs = Nothing

End Function

Private Function AdPt(str As String) As c_Coord
'collecte les point X, Y et Z de la string passée en argument
'format de la string : "vertex -3.954908e+000 2.330950e+000 1.093235e+000"
Dim col As Collection
Dim oCoor As c_Coord
Dim i As Long
Dim Valeur As Double

    Set col = SplitSpace(str)
    Set oCoor = New c_Coord
    For i = 2 To col.Count 'on sute la string "vertex"
        On Error Resume Next
        Valeur = CDbl(Replace(col.item(i), ".", ",", 1, 1, vbTextCompare))
        If err.Number = 0 Then
            oCoor.X = CDbl(Replace(col.item(i), ".", ",", 1, 1, vbTextCompare))
            oCoor.Y = CDbl(Replace(col.item(i + 1), ".", ",", 1, 1, vbTextCompare))
            oCoor.Z = CDbl(Replace(col.item(i + 2), ".", ",", 1, 1, vbTextCompare))
            Exit For
        Else
            err.Clear
        End If
        On Error GoTo 0
    Next i
    Set AdPt = oCoor
End Function

Public Function SplitSpace(str As String) As Collection
'Extrait les valeurs de la chaine str séparées par un espace
Dim oVal As Collection
Dim separateur As String

separateur = " "
    Set oVal = New Collection

    Do While InStr(1, str, separateur, vbTextCompare) > 0
        oVal.Add Left(str, InStr(1, str, separateur, vbTextCompare) - 1)
        str = Right(str, Len(str) - InStr(1, str, separateur, vbTextCompare))
    Loop
    oVal.Add str
    Set SplitSpace = oVal
    
    'Libération des objets
    Set oVal = Nothing
End Function
