

Module Module1



    Dim objDictionary2 As Scripting.Dictionary = CreateObject("Scripting.Dictionary")
    Dim intCantidadRenombrada As Integer = 0
    Dim intCantidadNoRenombrada As Integer = 0

    Function FileCount(localCurrentProduct As ProductStructureTypeLib.Product, localDictionary As Scripting.Dictionary)

        Dim i As Integer

        localCurrentProduct = localCurrentProduct.ReferenceProduct

        For i = 1 To localCurrentProduct.ReferenceProduct.Products.Count

            If localDictionary.Exists(localCurrentProduct.Products.Item(i).PartNumber) Then

                localDictionary.Item(localCurrentProduct.Products.Item(i).PartNumber) = localDictionary.Item(localCurrentProduct.Products.Item(i).PartNumber) + 1

                GoTo Finish

            ElseIf localCurrentProduct.Products.Item(i).PartNumber = localCurrentProduct.Products.Item(i).ReferenceProduct.Parent.Product.PartNumber Then

                localDictionary.Add(localCurrentProduct.Products.Item(i).PartNumber, 1)

                FileCount(localCurrentProduct.Products.Item(i), localDictionary)

            End If
Finish:
        Next

        FileCount = localDictionary.Count

    End Function




    '    Sub RenameInstances(ByRef objCurrentProduct As ProductStructureTypeLib.Product, localDictionary As Scripting.Dictionary)
    '        Dim i As Integer
    '        Dim j As Integer
    '        Dim k As Integer
    '        Dim arrRename(0) As String
    '        objCurrentProduct = objCurrentProduct.ReferenceProduct
    '        For i = 1 To objCurrentProduct.Products.Count
    '            ReDim Preserve arrRename(i)
    '            arrRename(i) = objCurrentProduct.Products.Item(i).PartNumber
    '            k = 0
    '            For j = 1 To i
    '                If arrRename(j) = objCurrentProduct.Products.Item(i).PartNumber Then
    '                    k += 1
    '                End If
    '            Next
    '            objCurrentProduct.Products.Item(i).Name = objCurrentProduct.Products.Item(i).PartNumber & "TEMP." & k
    '        Next
    '        For i = 1 To objCurrentProduct.Products.Count
    '            ReDim Preserve arrRename(i)
    '            arrRename(i) = objCurrentProduct.Products.Item(i).PartNumber
    '            k = 0
    '            For j = 1 To i
    '                If arrRename(j) = objCurrentProduct.Products.Item(i).PartNumber Then
    '                    k += 1
    '                End If
    '            Next
    '            objCurrentProduct.Products.Item(i).Name = objCurrentProduct.Products.Item(i).PartNumber & "." & k
    '            If localDictionary.Exists(objCurrentProduct.Products.Item(i).PartNumber) Then
    '                GoTo Finish
    '            ElseIf objCurrentProduct.Products.Item(i).PartNumber = objCurrentProduct.Products.Item(i).ReferenceProduct.Parent.Product.PartNumber Then
    '                localDictionary.Add(objCurrentProduct.Products.Item(i).PartNumber, 1)
    '                RenameInstances(objCurrentProduct.Products.Item(i), localDictionary)
    '            End If
    'Finish:

    '        Next
    '    End Sub



    ''' <summary>
    ''' Rutina principal para renombrar instancias.
    ''' Solo requiere el producto raíz.
    ''' </summary>
    Public Sub RenameInstances(objRootProduct As ProductStructureTypeLib.Product)
        Dim internalDict As Object = CreateObject("Scripting.Dictionary")
        ExecuteRename(objRootProduct, internalDict)

    End Sub

    ''' <summary>
    ''' Rutina interna que realiza el trabajo recursivo.
    ''' </summary>
    Private Sub ExecuteRename(ByVal currentProd As ProductStructureTypeLib.Product, ByVal localDict As Object)
        Dim refProduct As ProductStructureTypeLib.Product = currentProd.ReferenceProduct
        Dim children As ProductStructureTypeLib.Products = refProduct.Products
        Dim i As Integer, j As Integer, k As Integer

        For i = 1 To children.Count
            Dim child As ProductStructureTypeLib.Product = children.Item(i)
            k = 0
            For j = 1 To i
                If children.Item(j).PartNumber = child.PartNumber Then k += 1
            Next
            child.Name = child.PartNumber & "TEMP." & k
        Next


        For i = 1 To children.Count
            Dim child As ProductStructureTypeLib.Product = children.Item(i)
            k = 0
            For j = 1 To i
                If children.Item(j).PartNumber = child.PartNumber Then k += 1
            Next

            child.Name = child.PartNumber & "." & k

            If Not localDict.Exists(child.PartNumber) Then
                localDict.Add(child.PartNumber, 1)
                If child.Products.Count > 0 Then
                    ExecuteRename(child, localDict)
                End If
            End If
        Next
    End Sub





    Sub TextReplace(ByRef objCurrentProduct As ProductStructureTypeLib.Product, localDictionary As Scripting.Dictionary)

        Dim i As Integer

        ' Obtener los textos a buscar y reemplazar desde el formulario
        ' Dim strToSearch As String = Form1.TextBox1.Text
        ' Dim strReplacement As String = Form1.TextBox2.Text

        Dim strToSearch As String = "Texto a ser reemplazado"
        Dim strReplacement As String = "Texto de reemplazo"
        Dim strOldPartNumber As String


        objCurrentProduct = objCurrentProduct.ReferenceProduct

        For i = 1 To objCurrentProduct.Products.Count

            strOldPartNumber = objCurrentProduct.Products.Item(i).PartNumber

            If InStr(objCurrentProduct.Products.Item(i).PartNumber, strToSearch) <> 0 Then
                objCurrentProduct.Products.Item(i).PartNumber =
    Replace(Expression:=objCurrentProduct.Products.Item(i).PartNumber, Find:=strToSearch, Replacement:=strReplacement, 1, Count:=1, Compare:=1)
                If strOldPartNumber = objCurrentProduct.Products.Item(i).PartNumber Then
                    If objDictionary2.Exists(objCurrentProduct.Products.Item(i).PartNumber) Then
                        objDictionary2.Item(objCurrentProduct.Products.Item(i).PartNumber) = objDictionary2.Item(objCurrentProduct.Products.Item(i).PartNumber) + 1
                        GoTo Continuar
                    Else
                        objDictionary2.Add(objCurrentProduct.Products.Item(i).PartNumber, 1)
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
