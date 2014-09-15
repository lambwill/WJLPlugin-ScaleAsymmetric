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
        <CommandMethod("ScaleNonUniform", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub ScaleNonUniform() 'This command uses 'pickfirst' 
            Dim acSSPrompt As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection() 'If entities are selected use those, otherwise ask for selection

            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
            'These need to be user inputs...
            Dim origin = New Point3d(0, 0, 0)
            Dim scaleX As Double = 1
            Dim scaleY As Double = 2
            Dim scaleZ As Double = 1
            '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

            If (acSSPrompt.Status = PromptStatus.OK) Then
                ' There are selected entities...

                'Put the selected entities in a selection set
                Dim acSSet As SelectionSet = acSSPrompt.Value

                'Use the CreateAnonymousBlock function to add the entities to an anonymous block and get its ID
                Dim blkid As ObjectId = CreateAnonymousBlock(acSSet, origin)

                'Get the current drawing document (doc) and its database (db) and start a transaction (tr)
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim db As Database = doc.Database
                Using tr As Transaction = doc.TransactionManager.StartTransaction()
                    'Get block table record for the current space
                    Dim currSpace As BlockTableRecord = TryCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)

                    'Inset the anonymous block, add it to the btr and give it default properties
                    Dim insert As New BlockReference(origin, blkid)
                    currSpace.AppendEntity(insert)
                    insert.SetDatabaseDefaults()

                    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                    'Apply the block scaling
                    insert.ScaleFactors = New Scale3d(scaleX, scaleY, scaleZ)
                    '@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@

                    'tr.AddNewlyCreatedDBObject(insert, True)

                    'Explode the scaled block and add the constituent entities to a DBObjectCollection (acDBObjColl)
                    Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()
                    insert.Explode(acDBObjColl)

                    '' Add each entity from the exploded block to the block table record and the transaction
                    For Each acEnt As Entity In acDBObjColl
                        currSpace.AppendEntity(acEnt)
                        tr.AddNewlyCreatedDBObject(acEnt, True)
                    Next

                    tr.Commit()
                End Using
            Else
                ' There are no selected entities
                MsgBox("!")
            End If
        End Sub


        Public Shared Function CreateAnonymousBlock(acSSet As SelectionSet, origin As Point3d) As ObjectId
            '' ^might make (origin) optional

            '' Set the block name to"*U" in order to make the block anonymous
            Dim blockname As String = "*U"

            Dim blkid As ObjectId = ObjectId.Null

            '' Get the ??? WorkingDatabase ??? and start a transaction
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForWrite, False), BlockTable)

                '' Set up the new block
                Dim btr As New BlockTableRecord()
                btr.Name = blockname
                btr.Explodable = True
                btr.Origin = origin
                btr.BlockScaling = BlockScaling.Any


                '' Step through the objects in the selection set
                For Each acSSObj As SelectedObject In acSSet
                    '' Check to make sure a valid SelectedObject object was returned
                    If Not IsDBNull(acSSObj) Then
                        '' Open the selected object for read
                        Dim acEnt As Entity = tr.GetObject(acSSObj.ObjectId, OpenMode.ForRead)

                        If Not IsDBNull(acEnt) Then
                            '' If the entity exists, copy it into the block
                            btr.AppendEntity(acEnt.Clone)
                        End If
                    End If
                Next

                blkid = bt.Add(btr)
                tr.AddNewlyCreatedDBObject(btr, True)

                tr.Commit()

            End Using

            Return blkid '' Return the ObjectID of the anaonymous block
        End Function
    End Class

End Namespace