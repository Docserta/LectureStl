Attribute VB_Name = "TraceMailles"
Option Explicit

Public Sub CreateMailles(oVertexs)
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
Dim mHBody As HybridBody
Dim HSLinePTPT1 As HybridShapeLinePtPt, HSLinePTPT2 As HybridShapeLinePtPt, HSLinePTPT3 As HybridShapeLinePtPt
Dim HSFil As HybridShapeFill
Dim oVertex As c_Vertex
Dim NoPart As Integer 'decoupage du remontage en plusieurs part pour alléger les parts
Dim cptItem As Long 'Compteur d'items pour découpage des parts
Dim noVertex As Long
Dim mBar As c_ProgressBar

    'Initialisation des classes
    Set oVertex = New c_Vertex
    Set mDocs = CATIA.Documents
    Set mPartDoc = mDocs.Add("Part")
    Set mPart = mPartDoc.Part
    Set mHSFact = mPart.HybridShapeFactory
    Set HBodies = mPart.HybridBodies
    Set mHBody = HBodies.Add()
    Set mBar = New c_ProgressBar
        mBar.Affiche
        mBar.Titre = "Construction des surfaces"
    NoPart = 1
        
    On Error Resume Next
    For Each oVertex In oVertexs.Items
        noVertex = noVertex + 1
        cptItem = cptItem + 1
        'mBar.Balayage = cptItem
        
        'Discrimine le vertex si les points sont trop proches d'une droite
        If DiscrVertex(oVertex) Then
            'Creation du premier point
            Set mHSPtCoord1 = mHSFact.AddNewPointCoord(oVertex.PT1.X, oVertex.PT1.Y, oVertex.PT1.Z)
            mHBody.AppendHybridShape mHSPtCoord1
            mPart.InWorkObject = mHSPtCoord1
            'mPart.Update

            'Creation du second point
            Set mHSPtCoord2 = mHSFact.AddNewPointCoord(oVertex.Pt2.X, oVertex.Pt2.Y, oVertex.Pt2.Z)
            mHBody.AppendHybridShape mHSPtCoord2
            mPart.InWorkObject = mHSPtCoord2
            'mPart.Update

            'Creation du troisieme point
            Set mHSPtCoord3 = mHSFact.AddNewPointCoord(oVertex.Pt3.X, oVertex.Pt3.Y, oVertex.Pt3.Z)
            mHBody.AppendHybridShape mHSPtCoord3
            mPart.InWorkObject = mHSPtCoord3
            'mPart.Update

            'Création de la première droite
            Set HSLinePTPT1 = mHSFact.AddNewLinePtPt(mHSPtCoord1, mHSPtCoord2)
            mHBody.AppendHybridShape HSLinePTPT1
            mPart.InWorkObject = HSLinePTPT1
            'mPart.Update

            'Création de la seconde droite
            Set HSLinePTPT2 = mHSFact.AddNewLinePtPt(mHSPtCoord2, mHSPtCoord3)
            mHBody.AppendHybridShape HSLinePTPT2
            mPart.InWorkObject = HSLinePTPT2
            'mPart.Update

            'Création de la troisieme droite
            Set HSLinePTPT3 = mHSFact.AddNewLinePtPt(mHSPtCoord3, mHSPtCoord1)
            mHBody.AppendHybridShape HSLinePTPT3
            mPart.InWorkObject = HSLinePTPT3
            'mPart.Update

            'Création de la surface de rebouchage entre les trais ligne
            Set HSFil = mHSFact.AddNewFill()
            HSFil.AddBound HSLinePTPT1
            HSFil.AddBound HSLinePTPT2
            HSFil.AddBound HSLinePTPT3
            HSFil.Continuity = 0
            mHBody.AppendHybridShape HSFil
            mPart.InWorkObject = HSFil
            'mPart.Update

            If cptItem > NbItemDecoup Or noVertex = oVertexs.Count Then 'Création d'un nouveau Part
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
                Set mHBody = HBodies.Add()
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

Private Function DiscrVertex(oVertex As c_Vertex) As Boolean
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
                Set ptA = oVertex.PT1
                Set ptB = oVertex.Pt2
                Set ptC = oVertex.Pt3
            Case 2
                Set ptA = oVertex.Pt3
                Set ptB = oVertex.PT1
                Set ptC = oVertex.Pt2
            Case 3
                Set ptA = oVertex.Pt2
                Set ptB = oVertex.Pt3
                Set ptC = oVertex.PT1
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
