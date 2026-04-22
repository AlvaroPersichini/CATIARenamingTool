
Option Explicit On
Option Strict On


Module Module1

    Dim objDictionary2 As Scripting.Dictionary = CType(CreateObject("Scripting.Dictionary"), Scripting.Dictionary)
    Dim intCantidadRenombrada As Integer = 0
    Dim intCantidadNoRenombrada As Integer = 0

    Function FileCount(localCurrentProduct As ProductStructureTypeLib.Product, localDictionary As Scripting.Dictionary) As Integer
        Dim i As Object
        localCurrentProduct = localCurrentProduct.ReferenceProduct
        For i = 1 To localCurrentProduct.ReferenceProduct.Products.Count
            If localDictionary.Exists((localCurrentProduct.Products.Item(i).PartNumber)) Then
                localDictionary.Item(localCurrentProduct.Products.Item(i).PartNumber) = CType(localDictionary.Item(localCurrentProduct.Products.Item(i).PartNumber), Integer) + 1
                GoTo Finish
            ElseIf localCurrentProduct.Products.Item(i).PartNumber = CType(localCurrentProduct.Products.Item(i).ReferenceProduct.Parent, ProductStructureTypeLib.ProductDocument).Product.PartNumber Then
                localDictionary.Add(CType(localCurrentProduct.Products.Item(i).PartNumber, String), 1)
                FileCount(localCurrentProduct.Products.Item(i), localDictionary)
            End If
Finish:
        Next
        FileCount = localDictionary.Count
    End Function




    Public Sub RenameInstances(objRootProduct As ProductStructureTypeLib.Product)
        Dim internalDict As Object = CreateObject("Scripting.Dictionary")
        ExecuteRename(objRootProduct, internalDict)
    End Sub

    Private Sub ExecuteRename(ByVal currentProd As ProductStructureTypeLib.Product, ByVal localDict As Object)

        Dim refProduct As ProductStructureTypeLib.Product = currentProd.ReferenceProduct
        Dim children As ProductStructureTypeLib.Products = refProduct.Products
        Dim i As Integer, j As Integer, k As Integer
        Dim child As ProductStructureTypeLib.Product

        For i = 1 To children.Count
            child = children.Item(CType(i, Integer))
            k = 0
            For j = 1 To i
                If children.Item(CType(j, Integer)).PartNumber = child.PartNumber Then k += 1
            Next
            child.Name = child.PartNumber & "TEMP." & k
        Next

        For i = 1 To children.Count
            child = children.Item(CType(i, Integer))
            k = 0
            For j = 1 To i
                If children.Item(CType(j, Integer)).PartNumber = child.PartNumber Then k += 1
            Next
            child.Name = child.PartNumber & "." & k
            If Not CType(localDict, Scripting.Dictionary).Exists(CType(child.PartNumber, String)) Then
                CType(localDict, Scripting.Dictionary).Add(CType(child.PartNumber, String), 1)
                If child.Products.Count > 0 Then
                    ExecuteRename(child, localDict)
                End If
            End If
        Next

    End Sub



    Sub TextReplace(ByRef objCurrentProduct As ProductStructureTypeLib.Product, localDictionary As Scripting.Dictionary)

        Dim i As Integer
        Dim strToSearch As String = "Texto a ser reemplazado"
        Dim strReplacement As String = "Texto de reemplazo"
        Dim strOldPartNumber As String

        objCurrentProduct = objCurrentProduct.ReferenceProduct

        For i = 1 To objCurrentProduct.Products.Count
            strOldPartNumber = objCurrentProduct.Products.Item(CType(i, Object)).PartNumber
            If InStr(objCurrentProduct.Products.Item(CType(i, Object)).PartNumber, strToSearch) <> 0 Then
                objCurrentProduct.Products.Item(CType(i, Object)).PartNumber = Replace(Expression:=objCurrentProduct.Products.Item(CType(i, Object)).PartNumber, Find:=strToSearch, Replacement:=strReplacement, 1, Count:=1, Compare:=1)
                If strOldPartNumber = objCurrentProduct.Products.Item(CType(i, Object)).PartNumber Then
                    If objDictionary2.Exists(CType(objCurrentProduct.Products.Item(CType(i, Object)).PartNumber, String)) Then

                        objDictionary2.Item(objCurrentProduct.Products.Item(CType(i, Object)).PartNumber) = CType(objDictionary2.Item(objCurrentProduct.Products.Item(CType(i, Object)).PartNumber), Integer) + 1
                        GoTo Continuar
                    Else
                        objDictionary2.Add(CType(objCurrentProduct.Products.Item(CType(i, Object)).PartNumber, String), 1)
                        intCantidadNoRenombrada += 1
                    End If
                Else
                    intCantidadRenombrada += 1
                End If
Continuar:
            End If
        Next

        For i = 1 To objCurrentProduct.Products.Count
            If localDictionary.Exists(objCurrentProduct.Products.Item(i).PartNumber) Then
                localDictionary.Item(objCurrentProduct.Products.Item(i).PartNumber) = localDictionary.Item(objCurrentProduct.Products.Item(i).PartNumber) + 1
                GoTo Finish
            Else
                localDictionary.Add(objCurrentProduct.Products.Item(i).PartNumber, 1)
                TextReplace(objCurrentProduct.Products.Item(i), localDictionary)
            End If
Finish:
        Next

    End Sub

End Module
