Attribute VB_Name = "LectBinaire"
Public Function lectureSTLBinaire(FicSTL) As c_Vertexs
'Trace les triangles défini dans un fichier STL au format binaire
Dim oVertexs As c_Vertexs
Dim oVertex As c_Vertex
Dim Comment() As Variant
Dim NbTriangle As Integer
Dim NoTriangle As Long
Dim Normale As Single
Dim Coord(1 To 3) As Single
Dim tempCoord As c_Coord

Dim i As Long, CurOctet As Long

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
    Set tempCoord = New c_Coord
    
'charge le fichier slt
    'FicSTL = "C:\CFR\Dropbox\Macros\Lecture_STL\Bat52-part1-Export.STL"
    NoTriangle = 0
    Open FicSTL For Binary As #1

    'Récupère le commentaire (80 premiers octets)
    'For CurOctet = 1 To 80
        'Get #1, CurOctet, Comment
    'Next
    CurOctet = 80
    'Recupère le Nombre de triangle
    CurOctet = 81
        Get #1, CurOctet, NbTriangle
    CurOctet = Seek(1)
    
    'récupère les triangles
    Do While Not EOF(1)
        CurOctet = Seek(1) + 1
        NoTriangle = NoTriangle + 1
        'Récupère la normale au triangle '3*4 Octets
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1) + 1
        Next i
        
        'Récupère les 3 pts '3*3*4 Octets
        'Premier point
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1) + 1
        Next i
        tempCoord.X = Coord(1)
        tempCoord.Y = Coord(2)
        tempCoord.Z = Coord(3)
        oVertex.Pt1 = tempCoord
        
        'second point
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1) + 1
        Next i
        tempCoord.X = Coord(1)
        tempCoord.Y = Coord(2)
        tempCoord.Z = Coord(3)
        oVertex.Pt2 = tempCoord
        
        'Troisieme point
        For i = 1 To 3
            Get #1, CurOctet, Coord(i)
            CurOctet = Seek(1) + 1
        Next i
        tempCoord.X = Coord(1)
        tempCoord.Y = Coord(2)
        tempCoord.Z = Coord(3)
        oVertex.Pt3 = tempCoord

        oVertexs.Add oVertex.No, oVertex.Pt1, oVertex.Pt2, oVertex.Pt3
        
        'Passe les 2 octets de controle
        CurOctet = Seek(1) + 2
    Loop
  
err:
'Libération des objets
    Close #1
    
    Set lectureSTLBinaire = oVertexs
End Function

Private Function ouvreSTL() As String
'Recupere le fichier STL
'Dim NomComplet As String

    'Ouverture du fichier de paramètres
    ouvreSTL = CATIA.FileSelectionBox("Selectionner le fichier STL", "*.stl", CatFileSelectionModeOpen)
    If ouvreSTL = "" Then Exit Function 'on vérifie que qque chose a bien été selectionné
   
End Function

