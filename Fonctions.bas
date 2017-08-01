Attribute VB_Name = "Fonctions"
Option Explicit

Public Function VecteurDir(ptA As c_Coord, ptB As c_Coord) As c_Coord
'Calcule le vecteur directeur de 2 points
Dim tempVect As c_Coord
        
    Set tempVect = New c_Coord
        tempVect.X = ptB.X - ptA.X
        tempVect.Y = ptB.Y - ptA.Y
        tempVect.Z = ptB.Z - ptA.Z
    Set VecteurDir = tempVect
    
    'liberation des classes
    Set tempVect = Nothing
End Function

Public Function ProduitVect(VectA As c_Coord, VectB As c_Coord) As c_Coord
'Calcule le produit vectoriel de 2 vecteurs
Dim tempProd As c_Coord
    
    Set tempProd = New c_Coord
        tempProd.X = VectA.Y * VectB.Z - VectA.Z * VectB.Y
        tempProd.Y = VectA.Z * VectB.X - VectA.X * VectB.Z
        tempProd.Z = VectA.X * VectB.Y - VectA.Y * VectB.X
    Set ProduitVect = tempProd
    
    'liberation des classes
    Set tempProd = Nothing
End Function

Public Function NormVect(vect As c_Coord) As Double
'Calcule la norme d'un vecteur
    'NormVect = Sqr(Exp(vect.X) + Exp(vect.Y) + Exp(vect.Z))
    NormVect = Sqr(vect.X ^ 2 + vect.Y ^ 2 + vect.Z ^ 2)
End Function

Public Function DiscrVertex(oVertex As c_Vertex) As Boolean
'Discrimine le vertex si un des 3 points est trop proche de la droite opposé
Dim discrim As Boolean
Dim ptA As c_Coord
Dim ptB As c_Coord
Dim ptC As c_Coord
Dim AB As c_Coord
Dim BC As c_Coord
Dim ProdVectAB As c_Coord
Dim NormAB As Double
Dim NormBC As Double
Dim Distance As Double
Dim i As Integer

    DiscrVertex = True
    
    For i = 1 To 3
        Select Case i 'Rotation des points
            Case 1
                Set ptA = oVertex.Pt1
                Set ptB = oVertex.Pt2
                Set ptC = oVertex.Pt3
            Case 2
                Set ptA = oVertex.Pt3
                Set ptB = oVertex.Pt1
                Set ptC = oVertex.Pt2
            Case 3
                Set ptA = oVertex.Pt2
                Set ptB = oVertex.Pt3
                Set ptC = oVertex.Pt1
        End Select
        
        Set AB = VecteurDir(ptA, ptB)
        Set BC = VecteurDir(ptB, ptC)
        Set ProdVectAB = ProduitVect(AB, BC)
        NormAB = NormVect(ProdVectAB)
        NormBC = NormVect(BC)
        Distance = NormAB / NormBC
        If Distance < ValSeuil Then DiscrVertex = False
    Next i
    
    'Liberation des classes
    Set ptA = Nothing
    Set ptB = Nothing
    Set ptC = Nothing
    Set AB = Nothing
    Set BC = Nothing
    Set ProdVectAB = Nothing
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
