' (C) Copyright 2011 by  
'
Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(WJL_ScaleAsymmetric.MyCommands))> 
Namespace WJL_ScaleAsymmetric

    Public Class MyCommands

        ' Modal Command with pickfirst selection
        <CommandMethod("MyGroup", "MyPickFirst", "MyPickFirstLocal", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub MyPickFirst() ' This method can have any name
            Dim result As PromptSelectionResult = Application.DocumentManager.MdiActiveDocument.Editor.GetSelection()
            If (result.Status = PromptStatus.OK) Then
                ' There are selected entities
                ' Put your command using pickfirst set code here
            Else
                ' There are no selected entities
                ' Put your command code here
            End If
        End Sub

        <CommandMethod("ScaleObject")> _
        Public Sub ScaleObject()
            '' Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                '' Open the Block table for read
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                             OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                                OpenMode.ForWrite)

                '' Create a lightweight polyline
                Dim acPoly As Polyline = New Polyline()
                acPoly.AddVertexAt(0, New Point2d(1, 2), 0, 0, 0)
                acPoly.AddVertexAt(1, New Point2d(1, 3), 0, 0, 0)
                acPoly.AddVertexAt(2, New Point2d(2, 3), 0, 0, 0)
                acPoly.AddVertexAt(3, New Point2d(3, 3), 0, 0, 0)
                acPoly.AddVertexAt(4, New Point2d(4, 4), 0, 0, 0)
                acPoly.AddVertexAt(5, New Point2d(4, 2), 0, 0, 0)

                '' Close the polyline
                acPoly.Closed = True

                '' Reduce the object by a factor of 0.5 
                '' using a base point of (4,4.25,0)

                Dim pPtRes As PromptPointResult
                Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

                '' Prompt for the start point
                pPtOpts.Message = vbLf & "Enter scale origin: "
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                Dim origin As Point3d = pPtRes.Value

                Dim pDoubleRes As PromptDoubleResult

                Dim pDoubleOpts As PromptDoubleOptions = New PromptDoubleOptions("")

                '' Prompt for the X scale factor
                pDoubleOpts.Message = vbLf & "Enter X scale factor: "
                pDoubleRes = acDoc.Editor.GetDouble(pDoubleOpts)
                Dim scaleX As Double = pDoubleRes.Value

                '' Prompt for the Y scale factor
                pDoubleOpts.Message = vbLf & "Enter Y scale factor: "
                pDoubleRes = acDoc.Editor.GetDouble(pDoubleOpts)
                Dim scaleY As Double = pDoubleRes.Value

                '' Prompt for the Z scale factor
                pDoubleOpts.Message = vbLf & "Enter Z scale factor: "
                pDoubleRes = acDoc.Editor.GetDouble(pDoubleOpts)
                Dim scaleZ As Double = pDoubleRes.Value

                Dim matrixScale As Matrix3d

                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                'matrixScale = Matrix3d.Scaling(scaleAll, origin)


                'Dim scaleX As Double = scaleAll
                'Dim scaleY As Double = scaleAll
                'Dim scaleZ As Double = scaleAll

                Dim nulltransform As New Matrix3dBuilder

                nulltransform.ElementAt(0, 0) = scaleX
                nulltransform.ElementAt(0, 1) = 0
                nulltransform.ElementAt(0, 2) = 0
                nulltransform.ElementAt(0, 3) = origin.X * (-1 * (scaleX - 1))

                nulltransform.ElementAt(1, 0) = 0
                nulltransform.ElementAt(1, 1) = scaleY
                nulltransform.ElementAt(1, 2) = 0
                nulltransform.ElementAt(1, 3) = origin.Y * (-1 * (scaleY - 1))

                nulltransform.ElementAt(2, 0) = 0
                nulltransform.ElementAt(2, 1) = 0
                nulltransform.ElementAt(2, 2) = scaleZ
                nulltransform.ElementAt(2, 3) = origin.Z * (-1 * (scaleZ - 1))

                nulltransform.ElementAt(3, 0) = 0
                nulltransform.ElementAt(3, 1) = 0
                nulltransform.ElementAt(3, 2) = 0
                nulltransform.ElementAt(3, 3) = 1

                matrixScale = nulltransform.ToMatrix3d

                ''@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@
                acPoly.TransformBy(matrixScale)

                Dim myEd As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                myEd.WriteMessage(vbLf & "Scale by {0} at ({1},{2},{3}):" & vbLf & "{4}" & vbLf, scaleX, origin.X, origin.Y, origin.Z, matrixScale)


                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly)
                acTrans.AddNewlyCreatedDBObject(acPoly, True)

                '' Save the new objects to the database
                acTrans.Commit()
            End Using
        End Sub
        <CommandMethod("scales")> _
        Public Sub Scalestest()

            Dim origin As Point3d
            Dim scaleAll As Double

            'scale 1, origin (0,0,0):
            origin = New Point3d(0, 0, 0)
            scaleAll = 1
            MatrixTest(origin, scaleAll)

            'scale 2, origin (0,0,0):
            origin = New Point3d(0, 0, 0)
            scaleAll = 2
            MatrixTest(origin, scaleAll)

            'scale -1, origin (0,0,0):
            origin = New Point3d(0, 0, 0)
            scaleAll = -1
            MatrixTest(origin, scaleAll)

            'scale 5, origin (0,0,0):
            origin = New Point3d(0, 0, 0)
            scaleAll = 5
            MatrixTest(origin, scaleAll)

            'scale 1, origin (1,1,1):
            origin = New Point3d(1, 1, 1)
            scaleAll = 1
            MatrixTest(origin, scaleAll)

            'scale 1, origin (1,2,3):
            origin = New Point3d(1, 2, 3)
            scaleAll = 1
            MatrixTest(origin, scaleAll)

            'scale 1, origin (30,20,10):
            origin = New Point3d(30, 20, 10)
            scaleAll = 1
            MatrixTest(origin, scaleAll)
        End Sub

        Sub MatrixTest(origin As Point3d, scaleAll As Double)
            Dim matrixScale As Matrix3d = Matrix3d.Scaling(scaleAll, origin)
            Dim myEd As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            myEd.WriteMessage(vbLf & "Scale by {0} at ({1},{2},{3}):" & vbLf & "{4}" & vbLf, scaleAll, origin.X, origin.Y, origin.Z, matrixScale)
        End Sub
    End Class

End Namespace