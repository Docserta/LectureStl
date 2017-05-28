Attribute VB_Name = "Import_STL"


Option Explicit

Sub catmain()
'Trace les triangles défini dans un fichier STL
Dim oVertexs As c_Vertexs
Dim isBinaire  As Boolean
Dim mbar As c_ProgressBar
'Initialisation des classes
    Set oVertexs = New c_Vertexs
    Set mbar = New c_ProgressBar
        mbar.ProgressTitre 1, "lecture du fichier STL"
        mbar.Affiche

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
        Set oVertexs = lectureSTLBinaire(FicSTL, mbar)
    Else
        Set oVertexs = ColSTL(FicSTL, mbar)
    End If
'Tracé des mailles
    CreateMailles oVertexs, mbar

'Libération des objets
    Set oVertexs = Nothing

End Sub

Private Function ColSTL(FicSTL, mbar As c_ProgressBar) As c_Vertexs
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

Public Function lectureSTLBinaire(FicSTL, mbar As c_ProgressBar) As c_Vertexs
'Trace les triangles défini dans un fichier STL au format binaire
Dim oVertexs As c_Vertexs
Dim oVertex As c_Vertex
'Ctructure du fichier STL
Dim comment As String * 80
Dim NbTriangle As Single
Dim Normale As Long
Dim Coord(1 To 3) As Single
Dim Verif(1 To 2) As Byte

Dim tempCoord As c_Coord
Dim NoTriangle As Long
Dim i As Long, CurOctet As Variant
Dim cptItem As Long, posCurs As Long

On Error GoTo err
'Les 80 premiers octets sont un commentaire.
'Les 4 octets suivants forment un entier sur 32 bits qui représente le nombre de triangles présents dans le fichier.
'Ensuite, pour chaque triangle, on a une description sur 50 octets qui se décompose comme suit :
'3 fois 4 octets, chaque paquet de 4 octets représentant un flottant :
'     les coordonnées (x,y,z) de la direction normale au triangle
'     cette information est importante si on veut un rendu réaliste de l’objet
'    (elle conditionne la façon dont le triangle reflèteles rayons lumineux),
'     mais est inutile pour nous dans le cadre de ce projet.
'3 paquets de 3 fois fois 4 octets, chaque groupe de 4 octets représentant un flottant
'    les coordonnées (x,y,z) de chacun des sommets du triangle.
'Deux octets représentant un octet de contrôle (inutile dans le cadre ce projet).

'Initialisation des classes
    Set oVertexs = New c_Vertexs
    Set oVertex = New c_Vertex
    
'charge le fichier slt
    'FicSTL = "C:\CFR\Dropbox\Macros\Lecture_STL\Bat52-part1-Export.STL"
    NoTriangle = 0
    Open FicSTL For Binary As #1
    CurOctet = 1

    'Récupère le commentaire (80 premiers octets)
    Get #1, CurOctet, comment
    CurOctet = Seek(1)

    'Recupère le Nombre de triangle
    Get #1, CurOctet, NbTriangle
    CurOctet = Seek(1)
    
    'récupère les triangles
    Do While Not EOF(1)
        'affichage de la barre de progression
        '1 fois tous les 500 vertex pour ne pas ralentir le programme
        cptItem = cptItem + 1
            If cptItem Mod 500 = 0 Then
                posCurs = posCurs + 1
                mbar.Balayage = posCurs
                mbar.Etape = "Mesh N° " & cptItem
            End If
        NoTriangle = NoTriangle + 1
        
        'Récupère la normale au triangle '3*4 Octets
        Set tempCoord = New c_Coord
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1)
        Next i
        
        'Récupère les 3 pts '3*3*4 Octets
        'Premier point
        Set tempCoord = New c_Coord
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1)
        Next i
        tempCoord.X = Coord(1)
        tempCoord.Y = Coord(2)
        tempCoord.Z = Coord(3)
        oVertex.Pt1 = tempCoord
        
        'second point
        Set tempCoord = New c_Coord
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1)
        Next i
        tempCoord.X = Coord(1)
        tempCoord.Y = Coord(2)
        tempCoord.Z = Coord(3)
        oVertex.Pt2 = tempCoord
        
        'Troisieme point
        Set tempCoord = New c_Coord
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1)
        Next i
        tempCoord.X = Coord(1)
        tempCoord.Y = Coord(2)
        tempCoord.Z = Coord(3)
        oVertex.Pt3 = tempCoord

        oVertexs.Add oVertex.No, oVertex.Pt1, oVertex.Pt2, oVertex.Pt3
        
        'Passe les 2 octets de controle
        For i = 1 To 2
            Get #1, CurOctet, Verif(i)
            CurOctet = Seek(1)
        Next i
        Debug.Print CurOctet
    Loop
  
err:
'Libération des objets
    Close #1
    
    Set lectureSTLBinaire = oVertexs
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



Public Sub CreateMailles(oVertexs, mbar As c_ProgressBar)
'trace les mailles composées de trois lignes definies par les points X, Y, Y
'de la collection des vertex
'Découpage en plusieurs parts pour alléger l'update des surfaces

Dim mDocs As Documents
Dim mPartDoc As PartDocument
Dim mProd As Product
Dim mPart As Part
Dim mHSFact As HybridShapeFactory
Dim mHSPtCoord1 As HybridShapePointCoord, mHSPtCoord2 As HybridShapePointCoord, mHSPtCoord3 As HybridShapePointCoord
Dim HBodies As HybridBodies
Dim mHBodyPT As HybridBody, mHBodyLine As HybridBody, mHBodyMesh As HybridBody
Dim HSLinePTPT1 As HybridShapeLinePtPt, HSLinePTPT2 As HybridShapeLinePtPt, HSLinePTPT3 As HybridShapeLinePtPt
Dim HSFil As HybridShapeFill
Dim toto As HybridShapePlane3Points
Dim oVertex As c_Vertex
Dim NoPart As Integer 'decoupage du remontage en plusieurs part pour alléger les parts
Dim cptItem As Long 'Compteur d'items pour découpage des parts
Dim posCurs As Long 'Compteur pour avancement de la progress bar
Dim noVertex As Long

    'Initialisation des classes
    Set oVertex = New c_Vertex
    Set mDocs = CATIA.Documents
    Set mPartDoc = mDocs.Add("Part")
    Set mPart = mPartDoc.Part
    Set mHSFact = mPart.HybridShapeFactory
    Set HBodies = mPart.HybridBodies
    Set mHBodyPT = HBodies.Add()
    mHBodyPT.Name = "Points"
    Set mHBodyLine = HBodies.Add()
    mHBodyLine.Name = "Lines"
    Set mHBodyMesh = HBodies.Add()
    mHBodyMesh.Name = "Meshs"
    mbar.Titre = "Tracé des mailles"
    
    NoPart = 1
        
    On Error Resume Next
    For Each oVertex In oVertexs.Items
        noVertex = noVertex + 1
        cptItem = cptItem + 1
        
        'affichage de la barre de progression
        '1 fois tous les 500 vertex pour ne pas ralentir le programme
            If cptItem Mod 500 = 0 Then
                posCurs = posCurs + 1
                mbar.Balayage = posCurs
                mbar.Etape = noVertex
            End If
        
        'Discrimine le vertex si les points sont trop proches d'une droite
        If DiscrVertex(oVertex) Then
            'Creation du premier point
            Set mHSPtCoord1 = mHSFact.AddNewPointCoord(oVertex.Pt1.X, oVertex.Pt1.Y, oVertex.Pt1.Z)
            mHBodyPT.AppendHybridShape mHSPtCoord1
            mPart.InWorkObject = mHSPtCoord1
            'mPart.Update

            'Creation du second point
            Set mHSPtCoord2 = mHSFact.AddNewPointCoord(oVertex.Pt2.X, oVertex.Pt2.Y, oVertex.Pt2.Z)
            mHBodyPT.AppendHybridShape mHSPtCoord2
            mPart.InWorkObject = mHSPtCoord2
            'mPart.Update

            'Creation du troisieme point
            Set mHSPtCoord3 = mHSFact.AddNewPointCoord(oVertex.Pt3.X, oVertex.Pt3.Y, oVertex.Pt3.Z)
            mHBodyPT.AppendHybridShape mHSPtCoord3
            mPart.InWorkObject = mHSPtCoord3
            'mPart.Update

            'Création de la première droite
            Set HSLinePTPT1 = mHSFact.AddNewLinePtPt(mHSPtCoord1, mHSPtCoord2)
            mHBodyLine.AppendHybridShape HSLinePTPT1
            mPart.InWorkObject = HSLinePTPT1
            'mPart.Update

            'Création de la seconde droite
            Set HSLinePTPT2 = mHSFact.AddNewLinePtPt(mHSPtCoord2, mHSPtCoord3)
            mHBodyLine.AppendHybridShape HSLinePTPT2
            mPart.InWorkObject = HSLinePTPT2
            'mPart.Update

            'Création de la troisieme droite
            Set HSLinePTPT3 = mHSFact.AddNewLinePtPt(mHSPtCoord3, mHSPtCoord1)
            mHBodyLine.AppendHybridShape HSLinePTPT3
            mPart.InWorkObject = HSLinePTPT3
            'mPart.Update

            'Création de la surface de rebouchage entre les trais ligne
            Set HSFil = mHSFact.AddNewFill()
            HSFil.AddBound HSLinePTPT1
            HSFil.AddBound HSLinePTPT2
            HSFil.AddBound HSLinePTPT3
            HSFil.Continuity = 0
            mHBodyMesh.AppendHybridShape HSFil
            mPart.InWorkObject = HSFil
            'mPart.Update
            
            If cptItem > NbItemDecoup And DecoupFic Then 'Création d'un nouveau Part
                On Error GoTo 0
                cptItem = 1
                Set mProd = mPartDoc.Product
                mProd.PartNumber = "RemontageSTLPart" & NoPart
                mPartDoc.SaveAs "c:\temp\RemontSTLpart" & NoPart & ".Catpart"
                mPartDoc.Close
                NoPart = NoPart + 1
                Set mPartDoc = mDocs.Add("Part")
                Set mPart = mPartDoc.Part
                Set mHSFact = mPart.HybridShapeFactory
                Set HBodies = mPart.HybridBodies
                Set mHBodyPT = HBodies.Add()
                mHBodyPT.Name = "Points"
                Set mHBodyLine = HBodies.Add()
                mHBodyLine.Name = "Lines"
                Set mHBodyMesh = HBodies.Add()
                mHBodyMesh.Name = "Meshs"
                On Error Resume Next
            End If
        End If
    Next
    
    'Sauvegarde le dernier fichier
    Set mProd = mPartDoc.Product
    mProd.PartNumber = "RemontageSTLPart" & NoPart
    mPartDoc.SaveAs "c:\temp\RemontSTLpart" & NoPart & ".Catpart"
    mPartDoc.Close

End Sub


