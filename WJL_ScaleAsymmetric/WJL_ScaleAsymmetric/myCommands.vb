' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
'Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(WJL_ScaleAsymmetric.MyCommands))> 
Namespace WJL_ScaleAsymmetric

    Public Class MyCommands

        ' Modal Command with pickfirst selection
        <CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub MyPickFirst() ' This method can have any name
            Dim acSSPrompt As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection()
            Dim acSSet As SelectionSet

            If (acSSPrompt.Status = PromptStatus.OK) Then
                ' There are selected entities
                ' Put your command using pickfirst set code here
                acSSet = acSSPrompt.Value
            Else
                ' There are no selected entities
                ' Put your command code here
                MsgBox("!")

            End If

            Dim origin = New Point3d(0, 0, 0)
            Dim scaleX As Double = 1
            Dim scaleY As Double = 2
            Dim scaleZ As Double = 1

            Dim blkid As ObjectId = CreateAnonymousBlock(acSSet, origin)

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Using tr As Transaction = doc.TransactionManager.StartTransaction()
                Dim currSpace As BlockTableRecord = TryCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                Using insert As New BlockReference(origin, blkid)
                    currSpace.AppendEntity(insert)
                    insert.SetDatabaseDefaults()
                    insert.ScaleFactors = New Scale3d(scaleX, scaleY, scaleZ)
                    tr.AddNewlyCreatedDBObject(insert, True)
                    Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()
                    insert.Explode(acDBObjColl)

                    For Each acEnt As Entity In acDBObjColl
                        '' Add the new object to the block table record and the transaction
                        currSpace.AppendEntity(acEnt)
                        tr.AddNewlyCreatedDBObject(acEnt, True)
                    Next
                End Using
                tr.Commit()
            End Using

        End Sub


        Public Shared Function CreateAnonymousBlock(acSSet As SelectionSet, origin As Point3d) As ObjectId
            Dim blockname As String = "*U"

            Dim blkid As ObjectId = ObjectId.Null
            'Try

            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForWrite, False), BlockTable)

                Dim btr As New BlockTableRecord()
                btr.Name = blockname
                btr.Explodable = True
                btr.Origin = origin
                btr.BlockScaling = BlockScaling.Any


                '' Step through the objects in the selection set
                For Each acSSObj As SelectedObject In acSSet
                    '' Check to make sure a valid SelectedObject object was returned
                    If Not IsDBNull(acSSObj) Then
                        '' Open the selected object for write
                        Dim acEnt As Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)

                        If Not IsDBNull(acEnt) Then
                            '' Change the object's color to Green
                            btr.AppendEntity(acEnt.Clone)
                        End If
                    End If
                Next

                'For Each item As Entity In acSSet
                '    btr.AppendEntity(item)
                'Next

                blkid = bt.Add(btr)
                tr.AddNewlyCreatedDBObject(btr, True)

                tr.Commit()

            End Using
            'Catch ex As Autodesk.AutoCAD.Runtime.Exception
            ' Application.ShowAlertDialog((("ERROR: " & vbLf) + ex.Message & vbLf & "SOURCE: ") + ex.StackTrace)
            ' End Try
            Return blkid
        End Function
    End Class

End Namespace