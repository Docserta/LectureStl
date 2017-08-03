Attribute VB_Name = "CleanMesh"
Option Explicit

Sub CATMain()

Dim mDocs As Documents
Dim mPartDocIsol As PartDocument
Dim mPartDocStl As PartDocument
Dim mPartStl As Part
Dim mPartIsol As Part
Dim mProd As Product

Dim mHBodies As HybridBodies
Dim mHBody As HybridBody
Dim mSelStl As Selection
Dim mSelIsol As Selection

'Dim specsAndGeomWindow1 As SpecsAndGeomWindow
'Dim specsAndGeomWindow2 As SpecsAndGeomWindow
'Dim windows1 As Windows

Dim mlistFicStl() As String
Dim mFicStl As String

Dim i As Long

    Set mDocs = CATIA.Documents
    
    'Collecte de la liste des fichiers de remontage STL
    mlistFicStl = ListeFilesStl
    
    For i = 0 To UBound(mlistFicStl)
        mFicStl = PathFicStl & mlistFicStl(i)
    
        'Ouverture du part Stl et mise à jour
        Set mPartDocStl = mDocs.Open(mFicStl)
        Set mPartStl = mPartDocStl.Part
        mPartStl.Update
    
        'Création du part pour les ref isolées
        Set mPartDocIsol = mDocs.Add("Part")
    
        'Selection et copie du Set "Meshs" dans le part de remontage Stl
        Set mSelStl = mPartDocStl.Selection
        mSelStl.Clear
        Set mPartStl = mPartDocStl.Part
        Set mHBodies = mPartStl.HybridBodies
        Set mHBody = mHBodies.item("Meshs")
        mSelStl.Add mHBody
        mSelStl.Copy
        
        'Collage du set "Mesh" en tant que résultat dans le part Remontage STL isolé
        Set mSelIsol = mPartDocIsol.Selection
        mSelIsol.Clear
        Set mPartIsol = mPartDocIsol.Part
        mSelIsol.Add mPartIsol
        mSelIsol.PasteSpecial "CATPrtResultWithOutLink"
    
        'Fermeture du part STL
        mPartDocStl.Close
    
        'Sauvegarde du part isolé
        Set mProd = mPartDocIsol.Product
        mProd.PartNumber = RadNameFileIsol & i + 1
        mPartDocIsol.SaveAs "C:\temp\" & RadNameFileIsol & i + 1 & ".CATPart"
        mPartDocIsol.Close
    Next i
End Sub


