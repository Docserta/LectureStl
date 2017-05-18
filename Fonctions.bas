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
