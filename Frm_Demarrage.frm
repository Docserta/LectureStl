VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Frm_Demarrage 
   Caption         =   "Ajout d'une Grille de perçage"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   OleObjectBlob   =   "Frm_Demarrage.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Frm_Demarrage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False







Private Sub Btn_Navigateur_Click()
'Recupere le fichier STL

    'Ouverture du fichier de paramètres
    Me.TB_Fichier = CATIA.FileSelectionBox("Selectionner le fichier STL", "*.stl", CatFileSelectionModeOpen)

End Sub

Private Sub BtnAnnul_Click()
Me.Hide
Me.ChB_OkAnnule = False

End Sub

Private Sub BtnOK_Click()

Me.ChB_OkAnnule = True
Erreur = False

If Me.TB_Fichier = "" Then
    Me.ChB_OkAnnule = False
End If

If Not Erreur Then
    Me.Hide
End If
End Sub

Private Sub ChB_Decoup_Click()
    If Me.ChB_Decoup.Value = True Then
        Me.CBL_NbPt.Enabled = True
    Else
        Me.CBL_NbPt.Enabled = False
    End If
End Sub

Private Sub Logo_eXcent_Click()
'Chargement de la boite eXcent
    Load Frm_eXcent
    Frm_eXcent.Show
    Unload Frm_eXcent
End Sub



Private Sub UserForm_Initialize()
'ajoute les valeur de discrimination
    Me.CBL_Seuil.AddItem "0,02"
    Me.CBL_Seuil.AddItem "0,05"
    Me.CBL_Seuil.Value = "0,02"
'Ajoute le nombre de point de découpe des fichiers de resultat
    Me.CBL_NbPt.AddItem "5000"
    Me.CBL_NbPt.AddItem "10000"
    Me.CBL_NbPt.AddItem "15000"
    Me.CBL_NbPt.AddItem "20000"
    Me.CBL_NbPt.Value = "5000"
    Me.CBL_NbPt.Enabled = False
'Fichie ASCII par defaut
    Me.Rbt_ASCII = True
    
End Sub
