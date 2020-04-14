
'********************************************************************************
'This extension is Copyright (C) 2012-2016 Cyril Beaussier  - v. 1.4

'This code is free software; you can redistribute it and/or
'modify it under the terms of the CeCILL (for Ce[a] C[nrs] I[nria] L[ogiciel] L[ibre])
'License as published by CEA, CNRS and INRIA.

'This library is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
'******************************************************************************

Private oDoc, oDlg, oCurs as object
Private nPage as integer
Private bImage as boolean
Private uneImg as string
Private sLang as string
Private tPoids as double
Private nbImg as integer

Sub Main
    Dim oQuoi as object
    Dim oBibli, oDialog, oTaille As Object
    Dim oWait, oDlgModele, oBarre, oLocal As Object
    Dim sListe as string
    Dim sExt(1)
    Dim sImg, sPoids as string
    Dim nPoids as double
    Dim nHo, nLo, nHd, nLd, i as long

    oDoc = ThisComponent

    GlobalScope.BasicLibraries.LoadLibrary("Tools")
    DialogLibraries.LoadLibrary("PixCompress")
    oBibli = DialogLibraries.GetByName("PixCompress")
    oDialog = oBibli.GetByName("Dialogue")
    oDlg = CreateUnoDialog(oDialog)
    ' Language setting
    sLang = "en"
    oLocal = GetStarOfficeLocale()
    sLang = Left(oLocal.Language, 2)
    ' Dialog translation
    oDlg.Controls(0).Model.Label = Trans(2)
    oDlg.Controls(1).Model.Label = Trans(3)
    oDlg.Controls(3).Model.Label = Trans(1)
    oDlg.Controls(4).Model.Label = Trans(0)

    ' Test if the document has been saved
    If oDoc.isModified() or oDoc.URL = "" then
        msgBox Trans(10)
        exit Sub
    Endif

    bImage = false
    i = 0
    tPoids = 0
    oQuoi = oDoc.getCurrentSelection
    oImages = oDoc.getDrawPage()
    nbImg = oImages.Count -1
    oWait = CreeBarre(Trans(7), nbImg)
    oCurs = oDoc.CurrentController.ViewCursor
    nPage = oCurs.Page   ' Remember the current page


    listeImg = oDlg.GetControl("listeImg")
    oWait.setVisible( true )

    For Each oImg in oImages
      If oImg.supportsService("com.sun.star.text.TextGraphicObject") or _
         oImg.supportsService("com.sun.star.drawing.GraphicObjectShape") then
         oWait.Controls(0).Model.ProgressValue = i
         With oImg
             oGraph = .Graphic
               nHo = oGraph.Size100thMM.Height
            nLo = oGraph.Size100thMM.Width
            nHd = .Size.Height
            nLd = .Size.Width
            sExt = Split(.Graphic.MimeType, "/")
            sImg = Right(.GraphicUrl,32)
            sPoids = Poids(i)
            sListe = .Name & " (" & sExt(1)  & ") " & sPoids & Trans(8)
         End With
         listeImg.addItem (sListe, i)
         i = i + 1
      End If
    Next
    ' Display total file sizes
    nPoids = tPoids / 1000
    sPoids = Format(nPoids, "##,##0.00")
    oDlg.Controls(5).Model.Label = Trans(9) & sPoids & Trans(8)

    ' Is there an image selected?
    If oQuoi.ImplementationName = "SwXTextGraphicObject" or _
        oQuoi.ImplementationName = "SwXShape" then
        uneImg = oQuoi.name
        bImage = true
    Else
    ' List all the images
        bCompressOne = oDlg.GetControl("bCompressOne")
        bCompressOne.enable = false
    End if
    oWait.dispose()
    oDlg.execute
End Sub

Sub CompresseImages
    bImage = false
    Call CompresseImage
End Sub

Sub CompresseImage
    Dim oImages as object
    Dim sMsg as string
    Dim i as integer

    oImages = oDoc.getDrawPage()

    If not bImage then
    ' Make all the images
        oWait = CreeBarre(Trans(4), nbImg)
        oWait.setVisible( true )
        For Each oImg in oImages
            oWait.Controls(0).Model.ProgressValue = i
            Call CopieColle(oImg)
            i = i+1
        Next
        oWait.dispose()
        sMsg = Trans(5)
        oCurs.JumpToPage(nPage) ' Return the cursor to its initial position
    Else
    ' We only do the selected image
        For i = 0 To oImages.Count - 1
            oImg  = oImages(i)
            If  oImg.name = uneImg then exit for
        Next i
        Call CopieColle(oImg)
        sMsg = Trans(6) & oImg.name
    End if
    oDoc.store()
    ' Recalculate the size of the images
    sMsg = sMsg & chr(10) & Trans(11)
    sMsg = sMsg & Format(( tPoids / 1000 ), "##,##0.00") & Trans(8) & " > "
    tPoids = 0
    For i = 0 to nbImg
        sPoids = Poids(i)
    Next i
    sMsg = sMsg & Format(( tPoids / 1000 ), "##,##0.00") & Trans(8)
    ' Close the results dialog
    oDlg.EndExecute
    msgbox sMsg, 64, "PixCompress"
End Sub

Sub CopieColle (oImg)
    oPage = oDoc.CurrentController.Frame
    oDisp = createUnoService("com.sun.star.frame.DispatchHelper")
    oCopie = oDoc.createInstance("com.sun.star.drawing.GraphicObjectShape")

    Dim oTaille as new com.sun.star.awt.Size
    Dim nHo, nLo as long
    Dim nCoef as long
    nCoef = 1.5
    Dim args(0) as new com.sun.star.beans.PropertyValue
    args(0).Name = "SelectedFormat"
    args(0).Value = 2

    If oImg.supportsService("com.sun.star.text.TextGraphicObject") or _
        oImg.supportsService("com.sun.star.drawing.GraphicObjectShape") then

        oTaille = oImg.getSize()
        nHo = oTaille.Height
        nLo = oTaille.Width
        nHo = nHo * nCoef : nLo = nLo * nCoef
        oTaille.Height = nHo
        oTaille.Width = nLo

        oCopie.GraphicUrl = oImg.GraphicUrl
        oCopie.name = "_transfert_"
        oCopie.setSize( oTaille )
        oDoc.DrawPage.add(oCopie)

        oDoc.CurrentController.select(oCopie)
        oDisp.executeDispatch(oPage, ".uno:Cut", "", 0, Array())
        oDoc.CurrentController.select(oImg)
        oDisp.executeDispatch(oPage, ".uno:ClipboardFormatItems", "", 0, args())
    Endif
End Sub

Function Taille(sNom, nHo, nLo, nHd, nLd)
    nConv = 1000 ' mm > cm
    sTaille = "? x ?"
    If (nHo > 0) or (nLo > 0) then
        nHo = Format(nHo / nConv, "0.00")
        nLo = Format(nLo / nConv, "0.00")
        sTaille = nHo & " x " & nLo
    end if
    nHd = Format(nHd / nConv, "0.00")
    nLd = Format(nLd / nConv, "0.00")
    Taille = sNom & " : " & sTaille & " > " & nHd & " x " & nLd

End Function

Function Trans(nStr)
    Dim english, french, german, spanish

english = array ( _
    "Warning: the process cannot be canceled. Once did, original size for each picture will be lost.", _
    "Compression for :", _
    "All pictures", _
    "Selected picture", _
    "Compression in progress...", _
    "Compression finished for all pictures!", _
    "Compression finished for selected picture: ", _
    "Analyzis pictures", _
    "Kb", _
    "Weight total of pictures: ", _
    "Error: this document must be saved before!", _
    "Weight of pictures reduced: ", _
    )
french = array ( _
    "Attention : l'opération ne pourra être annulée. Une fois faite, la taille originale des images sera perdue.", _
    "Compression pour :", _
    "Toutes les images", _
    "L'image sélectionnée", _
    "Compression en cours...", _
    "Compression terminée pour toutes les images !", _
    "Compression terminée pour l'image sélectionnée : ", _
    "Analyse des images", _
    "Ko", _
    "Taille totale des images : ", _
    "Erreur : le document doit d'abord être sauvegardé !", _
    "Taille des images réduite : ", _
    )

    Select Case sLang
        Case "fr" : Trans = french(nStr)
        Case else : Trans = english(nStr)
    End select
End function


Function Poids(nIndex)
    Dim sRet, sUrl As string
    Dim args(1) As Variant
    Dim oZip As object, oImages As Object
    Dim oFlux As object, oImage As Object
    Dim nPoids as double
    Dim PropZ As New com.sun.star.beans.NamedValue

    sUrl = ThisComponent.URL
    oZip = createUnoService("com.sun.star.packages.Package")
    PropZ.Name = "RepairPackage"
    PropZ.Value = true
    Args(0) = sUrl
    Args(1) = PropZ
    oZip.initialize(Args())
    sRep = "Pictures"
    if oZip.hasByHierarchicalName(sRep) then
        oRep = oZip.getByHierarchicalName(sRep)
        oImages = oRep.getElementNames()
        For i = 0 to UBound(oImages)
            If i = nIndex Then
                oImage = oZip.getByHierarchicalName(sRep &"/"& oImages(i))
                oFlux = oImage.getInputStream()
                nPoids = oFlux.available()
                tPoids = tPoids + nPoids
                nPoids = nPoids / 1000
                sRet = Format(nPoids, "##,##0.00")
                oFlux.closeInput
            Endif
        Next i
    else
        sRet = "Error"
    endif
    Poids = sRet
End function

Function CreeBarre(sTitre as string, vMax as integer) as object
    ' Create the model
    oDlgModele = createUnoService( "com.sun.star.awt.UnoControlDialogModel")

    ' Position of the dialog
    oDlgModele.PositionX = 100
    oDlgModele.PositionY = 100
    oDlgModele.Width = 105
    oDlgModele.Height = 20
    oDlgModele.Title = sTitre

    ' Create the progress bar control
    oBarre = oDlgModele.createInstance( "com.sun.star.awt.UnoControlProgressBarModel" )
    ' Position of the progress bar
    oBarre.PositionX = 0
    oBarre.PositionY = 0
    oBarre.Width = 105
    oBarre.Height = 20
    oBarre.ProgressValueMin = 0
    oBarre.ProgressValueMax = vMax
    ' Put the bar into the model
    oDlgModele.insertByName("bar", oBarre )

    ' Create a dialog from this model
    oWait = createUnoService( "com.sun.star.awt.UnoControlDialog")
    oWait.setModel( oDlgModele )

    CreeBarre = oWait
End function