Attribute VB_Name = "Nettoyage"
Option Explicit


Sub CATMain()

Dim mDocs As Documents
Dim mPartDoc As PartDocument
Dim mPart As Part
Dim HBodies As HybridBodies
Dim mHBody As HybridBody
Dim mHShapes As HybridShapes
Dim mHShape As HybridShape
Dim mSelection As Selection


    'initialisation des variables
    Set mPartDoc = CATIA.ActiveDocument
    Set mPart = mPartDoc.Part
    Set HBodies = mPart.HybridBodies
    Set mHBody = HBodies.item(1)
    'Set mHBody = HBodies.item("Geometrical Set.1")
    Set mSelection = mPartDoc.Selection
    Set mHShapes = mHBody.HybridShapes
    
    For Each mHShape In mHShapes
        On Error GoTo err_update
        mSelection.Clear
        mPart.UpdateObject mHShape
        GoTo suite
err_update:

            mSelection.Add mHShape
            mSelection.Delete

suite:
    Next

End Sub
