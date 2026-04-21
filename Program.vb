Option Explicit On
Option Strict On
Imports CATIAClassLibrary


Module Program


    Sub Main()

        ' Inicio
        Console.WriteLine(">>> Starting Process...")



        ' Catia
        Dim CATIAsession As New CatiaSession
        If Not CATIAsession.IsReady Then
            MsgBox(CATIAsession.Description)
            Exit Sub
        End If
        Dim oProduct As ProductStructureTypeLib.Product = CATIAsession.RootProduct
        CATIAsession.Application.DisplayFileAlerts = False



        RenameInstances(oProduct)

        Console.WriteLine(">>> Process Finished!")

    End Sub



End Module
