
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System;
using System.Security.Cryptography.X509Certificates;
using System.IO;
using System.Xml.Linq;

namespace WorkWithBuildings
{
    public class SignBuildings
    {
        public static string buildSpecTextValue = null;
        public static PromptEntityResult buildSpecificationExample = null;
        public static PromptEntityResult buildDescriptionExample = null;
        public static PromptEntityResult buildNumberExample = null;
        public double RotateAngleFix(double angle)
        {
            double sin = Math.Sin(angle);
            double cos = Math.Cos(angle);
            if (sin < 0 && cos > 0)
                angle += 0;
            else if (sin > 0 && cos > 0)
                angle += 0;
            else if (sin < 0 && cos < 0)
                angle += Math.PI;
            else if (sin > 0 && cos < 0)
                angle -= Math.PI;
            else if (sin == 0)
                angle = 0;
            else
                angle = 0;
            return angle;
        }

        public Polyline3d StreetCreate()
        {
            Point3dCollection streetPoints = new Point3dCollection();
            Polyline3d streetPline = new Polyline3d();
            while (true)
            {
                PromptPointResult giveMePoint = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetPoint("\nУкажите точку на улице. Нажмите Esc если достаточно");
                if (giveMePoint.Status != PromptStatus.OK)
                    break;
                else
                {
                    Point3d streetPoint = giveMePoint.Value;
                    streetPoints.Add(streetPoint);
                }
            }
            streetPline = new Polyline3d(0, streetPoints, false);
            return streetPline;
        }

        public bool EntitysBoundIntersectCheck(Point3d centreOfEntity, Entity obj1, Entity obj2)
        {
            bool intersect = false;
            if (obj1.Bounds != null)
            {
                Point3d boundBoxMinPoint = obj1.Bounds.Value.MinPoint;
                Point3d boundBoxMaxPoint = obj1.Bounds.Value.MaxPoint;
                Polyline specBoundBox = new Polyline();
                specBoundBox.AddVertexAt(0, new Point2d((centreOfEntity.X + boundBoxMinPoint.X), (centreOfEntity.Y + boundBoxMinPoint.Y)), 0, 0, 0);
                specBoundBox.AddVertexAt(1, new Point2d((centreOfEntity.X + boundBoxMinPoint.X), (centreOfEntity.Y + boundBoxMaxPoint.Y)), 0, 0, 0);
                specBoundBox.AddVertexAt(2, new Point2d((centreOfEntity.X + boundBoxMaxPoint.X), (centreOfEntity.Y + boundBoxMaxPoint.Y)), 0, 0, 0);
                specBoundBox.AddVertexAt(3, new Point2d((centreOfEntity.X + boundBoxMaxPoint.X), (centreOfEntity.Y + boundBoxMinPoint.Y)), 0, 0, 0);
                specBoundBox.Closed = true;
                specBoundBox.Elevation = 0;
                Point3dCollection charIntersectContur = new Point3dCollection();
                specBoundBox.IntersectWith(obj2, Intersect.OnBothOperands, new Plane(), charIntersectContur, IntPtr.Zero, IntPtr.Zero);
                if (charIntersectContur.Count > 0)
                    intersect = true;
            }
            else
                intersect = false;

            return intersect;
        }


        [CommandMethod("SignBuildings")]
        [CommandMethod("ПодписатьДома")]
        public void SigningBuildings()     ///убираем блоки накладываемые на другие объекты
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database dB = doc.Database;
            Editor ed = doc.Editor;

            // Предлагаем пользователю выбрать примеры примитивов подписей
            if (buildSpecificationExample == null)
            {
                buildSpecificationExample = doc.Editor.GetEntity("\nВыберите пример подписи характеристики строения. Текст или Мтекст");

                if (buildSpecificationExample.ObjectId == ObjectId.Null)
                {
                    doc.Editor.WriteMessage("\nНичего не выбрано. Программа прекратила работу");
                    return;
                }
                else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName != "TEXT" & buildSpecificationExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                {
                    while (buildSpecificationExample.ObjectId.ObjectClass.DxfName != "TEXT" & buildSpecificationExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                    {
                        buildSpecificationExample = doc.Editor.GetEntity("\nВыберите пример подписи характеристики строения. Текст или Мтекст");
                    }
                }
            }
            if (buildDescriptionExample == null)
            {
                buildDescriptionExample = doc.Editor.GetEntity("\nВыберите пример подписи описания строения. Текст или Мтекст");
                if (buildDescriptionExample.ObjectId == ObjectId.Null)
                {
                    doc.Editor.WriteMessage("\nНичего не выбрано. Программа прекратила работу");
                    return;
                }
                else if (buildDescriptionExample.ObjectId.ObjectClass.DxfName != "TEXT" & buildDescriptionExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                {
                    while (buildDescriptionExample.ObjectId.ObjectClass.DxfName != "TEXT" & buildDescriptionExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                    {
                        buildDescriptionExample = doc.Editor.GetEntity("\nВыберите пример подписи описания строения. Текст или Мтекст");
                    }
                }
            }

            if (buildNumberExample == null)
            {
                buildNumberExample = doc.Editor.GetEntity("\nВыберите пример подписи номера дома. Текст или Мтекст");
                if (buildNumberExample.ObjectId == ObjectId.Null)
                {
                    doc.Editor.WriteMessage("\nНичего не выбрано. Программа прекратила работу");
                    return;
                }
                else if (buildNumberExample.ObjectId.ObjectClass.DxfName != "TEXT" & buildNumberExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                {
                    while (buildNumberExample.ObjectId.ObjectClass.DxfName != "TEXT" & buildNumberExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                    {
                        buildNumberExample = doc.Editor.GetEntity("\nВыберите пример подписи номера дома. Текст или Мтекст");
                    }
                }
            }
            Polyline3d street = StreetCreate();

            while (true)
            {
                // Starts a new transaction with the Transaction Manager
                using (Transaction trans = dB.TransactionManager.StartTransaction())
                {
                    // Получение доступа к активному пространству (пространство модели или лист)
                    BlockTableRecord currentSpace = trans.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                    PromptEntityResult buildingSelection = doc.Editor.GetEntity("\nВыберите контур здания");
                    if (buildingSelection.ObjectId == ObjectId.Null)
                        break;
                    else if (buildingSelection.ObjectId.ObjectClass.DxfName != "POLYLINE" && buildingSelection.ObjectId.ObjectClass.DxfName != "LWPOLYLINE")
                        continue;
                    else
                    {
                        //открываем для чтения
                        DBObject buildingСontour = buildingSelection.ObjectId.GetObject(OpenMode.ForWrite);
                        Polyline buildingPline = buildingСontour as Polyline;
                        Point3d centreOfBuilding = new Point3d(0, 0, 0);
                        buildingPline.Elevation = 0;

                        List<Point2d> buildingsPoints2d = new List<Point2d>();
                        for (int i = 0; i < buildingPline.NumberOfVertices; i++)
                            buildingsPoints2d.Add(buildingPline.GetPoint2dAt(i));
                        double distBetweenVertexesOfBuilding = 0;
                        for (int i = 0; i < buildingsPoints2d.Count; i++)
                        {
                            int nextVertexIndex = i < buildingsPoints2d.Count - 1 ? i + 1 : 0;
                            distBetweenVertexesOfBuilding = buildingsPoints2d[i].GetDistanceTo(buildingsPoints2d[nextVertexIndex]);
                            if (distBetweenVertexesOfBuilding < 0.1)
                            {
                                buildingsPoints2d.RemoveAt(i);
                                i -= 1;
                            }
                        }

                        //высчитываем центр
                        if (buildingsPoints2d.Count < 3)
                        {
                            doc.Editor.WriteMessage("\nВыберите полилинию в которой 3 вершины или больше");
                            continue;
                        }
                        else if (buildingsPoints2d.Count == 4)
                        {
                            for (int i = 0; i < buildingsPoints2d.Count; i++)
                            {
                                centreOfBuilding += buildingPline.GetPoint3dAt(i).GetAsVector();
                            }
                            centreOfBuilding /= buildingsPoints2d.Count;
                            centreOfBuilding = new Point3d(centreOfBuilding.X, centreOfBuilding.Y, 0);
                        }
                        else if (buildingsPoints2d.Count != 4)
                        {
                            //достаем значение площади                          
                            double buildingArea = buildingPline.Area;

                            double numToMultiply = 1 / (6 * buildingArea);
                            double coordX = 0;
                            double coordY = 0;
                            for (int i = 0; i < buildingsPoints2d.Count; i++)
                            {
                                coordX += i < buildingsPoints2d.Count - 1 ? ((((buildingPline.GetPoint3dAt(i).X + buildingPline.GetPoint3dAt(i + 1).X) * ((buildingPline.GetPoint3dAt(i).X * buildingPline.GetPoint3dAt(i + 1).Y) - (buildingPline.GetPoint3dAt(i + 1).X * buildingPline.GetPoint3dAt(i).Y))))) :
                                    (((buildingPline.GetPoint3dAt(i).X + buildingPline.GetPoint3dAt(0).X) * ((buildingPline.GetPoint3dAt(i).X * buildingPline.GetPoint3dAt(0).Y) - (buildingPline.GetPoint3dAt(0).X * buildingPline.GetPoint3dAt(i).Y))));
                                coordY += i < buildingsPoints2d.Count - 1 ? (((buildingPline.GetPoint3dAt(i).Y + buildingPline.GetPoint3dAt(i + 1).Y) * ((buildingPline.GetPoint3dAt(i).X * buildingPline.GetPoint3dAt(i + 1).Y) - (buildingPline.GetPoint3dAt(i + 1).X * buildingPline.GetPoint3dAt(i).Y)))) :
                                    (((buildingPline.GetPoint3dAt(i).Y + buildingPline.GetPoint3dAt(0).Y) * ((buildingPline.GetPoint3dAt(i).X * buildingPline.GetPoint3dAt(0).Y) - (buildingPline.GetPoint3dAt(0).X * buildingPline.GetPoint3dAt(i).Y))));
                            }
                            if (coordX < 0 || coordY < 0)
                            {
                                coordX *= -1;
                                coordY *= -1;
                            }
                            centreOfBuilding = new Point3d(coordX, coordY, 0) * numToMultiply;
                        }

                        // зная координаты центра и 2 точки лежащие по краям улицы определяем сегмент полилинии вдоль которого поворачиваем текст
                        // 1: вычисляем координаты точки пересечения улицы и перпендикуляра опущенного из центра строения к "улице"
                        Point3d fromBuildToStreet = street.GetClosestPointTo(centreOfBuilding, Vector3d.ZAxis, false);
                        fromBuildToStreet = new Point3d(fromBuildToStreet.X, fromBuildToStreet.Y, 0);

                        // 2: Определяем сегмент полилинии вдоль которого нужно разворачивать подпись                        
                        LineSegment2d segmentForRotate = new LineSegment2d();
                        LineSegment2d lineFromCentreToStreet = new LineSegment2d(new Point2d(fromBuildToStreet.X, fromBuildToStreet.Y), new Point2d(centreOfBuilding.X, centreOfBuilding.Y));
                        Point2d[] pointOnSideToRotateCollection = new Point2d[1];

                        for (int i = 0; i < buildingsPoints2d.Count; i++)
                        {
                            int indexOfNextVertex = i < buildingsPoints2d.Count - 1 ? i + 1 : 0;
                            LineSegment2d buildingPlineSegment = new LineSegment2d(buildingsPoints2d[i], buildingsPoints2d[indexOfNextVertex]);
                            pointOnSideToRotateCollection = lineFromCentreToStreet.IntersectWith(buildingPlineSegment);
                            if (pointOnSideToRotateCollection != null)
                            {
                                segmentForRotate = buildingPlineSegment;
                                break;
                            }
                        }

                        double segmentForRotateLength = segmentForRotate.Length;

                        Vector3d vectorFromBuiding = centreOfBuilding - fromBuildToStreet;
                        Vector3d vectorSegmentForRotate =
                            new Point3d(segmentForRotate.EndPoint.X, segmentForRotate.EndPoint.Y, 0)
                            - new Point3d(segmentForRotate.StartPoint.X, segmentForRotate.StartPoint.Y, 0);

                        double angelOfRotation = vectorSegmentForRotate.GetAngleTo(Vector3d.XAxis, -Vector3d.ZAxis);
                        angelOfRotation = RotateAngleFix(angelOfRotation);


                        DBText buildingSpecifications = new DBText();
                        DBText buildDescription = new DBText();
                        DBText buildingNumber = new DBText();
                        MText buildingSpecificationsMt = new MText();
                        MText buildDescriptionMt = new MText();
                        MText buildingNumberMt = new MText();
                        double buildingSpecificationsHeight = 0;
                        double buildingSpecificationsWeight = 0;

                        DBObject buildSpecDBObject = buildSpecificationExample.ObjectId.GetObject(OpenMode.ForRead);
                        DBObject buildDescriptionDBObject = buildDescriptionExample.ObjectId.GetObject(OpenMode.ForRead);
                        DBObject buildNumberDBObject = buildNumberExample.ObjectId.GetObject(OpenMode.ForRead);

                        DBText buildSpecText = new DBText();
                        MText buildSpecMText = new MText();
                        DBText buildDescriptionText = new DBText();
                        MText buildDescriptionMText = new MText();
                        DBText buildNumberText = new DBText();
                        MText buildNumberMText = new MText();

                        PromptStringOptions optBuildSpecText = new PromptStringOptions("\nВведите характеристики строения. Если не требуется нажмите клавишу Esc")
                        {
                            DefaultValue = buildSpecTextValue,
                            UseDefaultValue = true
                        };

                        PromptResult buildSpec = doc.Editor.GetString(optBuildSpecText);
                        if (buildSpec.StringResult == "")
                            buildSpecTextValue = null;
                        else
                            buildSpecTextValue = buildSpec.StringResult;

                        if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "TEXT")
                        {
                            buildSpecText = buildSpecDBObject as DBText;
                            buildingSpecifications.TextString = buildSpec.StringResult;
                            buildingSpecifications.Color = buildSpecText.Color;
                            buildingSpecifications.Height = buildSpecText.Height;
                            buildingSpecificationsHeight = buildSpecText.Height;
                            buildingSpecifications.Layer = buildSpecText.Layer;
                            buildingSpecifications.TextStyleId = buildSpecText.TextStyleId;
                            if (buildSpec.StringResult != "")
                                buildingSpecificationsWeight = buildingSpecifications.Bounds.Value.MaxPoint.X - buildingSpecifications.Bounds.Value.MinPoint.X;
                            buildingSpecifications.Rotation = angelOfRotation;
                            buildingSpecifications.HorizontalMode = TextHorizontalMode.TextMid; //если использовать что-то кроме mid или centre текст создается в точке с нулевыми координатами
                        }
                        else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                        {
                            buildSpecMText = buildSpecDBObject as MText;
                            buildingSpecificationsMt.Contents = buildSpec.StringResult;
                            buildingSpecificationsMt.Color = buildSpecMText.Color;
                            buildingSpecificationsMt.TextHeight = buildSpecMText.TextHeight;
                            buildingSpecificationsHeight = buildSpecMText.TextHeight;
                            buildingSpecificationsMt.Layer = buildSpecMText.Layer;
                            buildingSpecificationsMt.TextStyleId = buildSpecMText.TextStyleId;
                            if (buildSpec.StringResult != "")
                                buildingSpecificationsWeight = buildingSpecificationsMt.Bounds.Value.MaxPoint.X - buildingSpecificationsMt.Bounds.Value.MinPoint.X;
                            buildingSpecificationsMt.Rotation = angelOfRotation;
                            buildingSpecificationsMt.Attachment = AttachmentPoint.MiddleCenter;
                        }

                        PromptStringOptions buildDescOptions = new PromptStringOptions("\nВведите описание,если требуется.Ввод оканчивается нажатием клавиши Enter");
                        buildDescOptions.AllowSpaces = true;
                        PromptResult buildDesc = doc.Editor.GetString(buildDescOptions);
                        double buildDescriptionWidth = 0;
                        if (buildDesc.StringResult == "")
                        {
                            if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                if (buildingSpecificationsWeight > segmentForRotateLength - 0.5)
                                    buildingSpecifications.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));

                                buildingSpecifications.AlignmentPoint = centreOfBuilding;

                                bool specIntersect = EntitysBoundIntersectCheck(centreOfBuilding, buildingSpecifications, buildingPline);
                                if (specIntersect == true)
                                    buildingSpecifications.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));

                                currentSpace.AppendEntity(buildingSpecifications);
                                trans.AddNewlyCreatedDBObject(buildingSpecifications, true);
                            }
                            else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                            {
                                if (buildingSpecificationsWeight > segmentForRotateLength - 0.5)
                                    buildingSpecificationsMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));

                                buildingSpecificationsMt.Location = centreOfBuilding;

                                bool specIntersect = EntitysBoundIntersectCheck(centreOfBuilding, buildingSpecificationsMt, buildingPline);
                                if (specIntersect == true)
                                    buildingSpecificationsMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));

                                currentSpace.AppendEntity(buildingSpecificationsMt);
                                trans.AddNewlyCreatedDBObject(buildingSpecificationsMt, true);
                            }
                        }
                        else
                        {
                            double centreOfBuildingSpecificationsX = Math.Abs(((centreOfBuilding[0] - centreOfBuilding[0]) * Math.Cos(angelOfRotation)) - (((centreOfBuilding[1] + (0.65 * buildingSpecificationsHeight)) - centreOfBuilding[1]) * (Math.Sin(angelOfRotation) * -1) + centreOfBuilding[0]));
                            double centreOfBuildingSpecificationsY = Math.Abs(((centreOfBuilding[0] - centreOfBuilding[0]) * (Math.Sin(angelOfRotation * -1))) + (((centreOfBuilding[1] + (0.65 * buildingSpecificationsHeight)) - centreOfBuilding[1]) * Math.Cos(angelOfRotation) + centreOfBuilding[1]));
                            double centreOfBuildDescriptionX = Math.Abs((((centreOfBuilding[0]) - centreOfBuilding[0]) * Math.Cos(angelOfRotation)) - (((centreOfBuilding[1] - (0.65 * buildingSpecificationsHeight)) - centreOfBuilding[1]) * (Math.Sin(angelOfRotation) * -1) + centreOfBuilding[0]));
                            double centreOfBuildDescriptionY = Math.Abs((((centreOfBuilding[0]) - centreOfBuilding[0]) * (Math.Sin(angelOfRotation) * -1)) + (((centreOfBuilding[1] - (0.65 * buildingSpecificationsHeight)) - centreOfBuilding[1]) * Math.Cos(angelOfRotation) + centreOfBuilding[1]));

                            Point3d centreOfBuildingSpecifications = new Point3d(centreOfBuildingSpecificationsX, centreOfBuildingSpecificationsY, 0);
                            Point3d centreOfBuildDescription = new Point3d(centreOfBuildDescriptionX, centreOfBuildDescriptionY, 0);
                            Point3d movedTextOrto = centreOfBuildingSpecifications.RotateBy(1.5707963267949, -Vector3d.ZAxis, centreOfBuilding);
                            Point3d movedText2Orto = centreOfBuildDescription.RotateBy(1.5707963267949, -Vector3d.ZAxis, centreOfBuilding);

                            Point3d centreOfBuildingSpecificationsOrto = new Point3d();
                            Point3d centreOfBuildDescriptionOrto = new Point3d();
                            if (movedTextOrto.Y > movedText2Orto.Y)
                            {
                                centreOfBuildingSpecificationsOrto = movedTextOrto;
                                centreOfBuildDescriptionOrto = movedText2Orto;
                            }
                            else
                            {
                                centreOfBuildingSpecificationsOrto = movedText2Orto;
                                centreOfBuildDescriptionOrto = movedTextOrto;
                            }

                            if (buildDescriptionExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                buildDescriptionText = buildDescriptionDBObject as DBText;
                                buildDescription.TextString = buildDesc.StringResult;
                                buildDescription.Color = buildDescriptionText.Color;
                                buildDescription.Height = buildDescriptionText.Height;
                                buildDescription.Layer = buildDescriptionText.Layer;
                                buildDescription.TextStyleId = buildDescriptionText.TextStyleId;
                                Extents3d buildDescriptionBound = (Extents3d)buildDescription.Bounds;
                                buildDescriptionWidth = buildDescriptionBound.MaxPoint[0] - buildDescriptionBound.MinPoint[0];
                                buildDescription.Rotation = angelOfRotation;
                                buildDescription.HorizontalMode = TextHorizontalMode.TextMid; //если использовать что-то кроме mid или centre текст создается в точке с нулевыми коорлдинатами
                            }
                            else if (buildDescriptionExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                            {
                                buildDescriptionMText = buildDescriptionDBObject as MText;
                                buildDescriptionMt.Contents = buildDesc.StringResult;
                                buildDescriptionMt.Color = buildDescriptionMText.Color;
                                buildDescriptionMt.TextHeight = buildDescriptionMText.TextHeight;
                                buildDescriptionMt.Layer = buildDescriptionMText.Layer;
                                buildDescriptionMt.TextStyleId = buildDescriptionMText.TextStyleId;
                                Extents3d buildDescriptionBound = (Extents3d)buildDescription.Bounds;
                                buildDescriptionWidth = buildDescriptionBound.MaxPoint[0] - buildDescriptionBound.MinPoint[0];
                                buildDescriptionMt.Rotation = angelOfRotation;
                                buildDescriptionMt.Attachment = AttachmentPoint.MiddleCenter;
                            }

                            if (buildSpecTextValue == null && buildDescriptionExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                if (buildDescriptionWidth > segmentForRotateLength - 1.0)
                                {
                                    buildDescription.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                }
                                buildDescription.AlignmentPoint = centreOfBuilding;
                                currentSpace.AppendEntity(buildDescription);
                                trans.AddNewlyCreatedDBObject(buildDescription, true);
                            }
                            else if (buildSpecTextValue == null && buildDescriptionExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                            {
                                if (buildDescriptionWidth > segmentForRotateLength - 1.0)
                                {
                                    buildDescriptionMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                }
                                buildingSpecificationsMt.Location = centreOfBuilding;
                                currentSpace.AppendEntity(buildDescriptionMt);
                                trans.AddNewlyCreatedDBObject(buildDescriptionMt, true);
                            }
                            else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "TEXT" && buildDescriptionExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                if (buildDescriptionWidth < segmentForRotateLength - 1.0)
                                {
                                    buildingSpecifications.AlignmentPoint = centreOfBuildingSpecifications;
                                    buildDescription.AlignmentPoint = centreOfBuildDescription;
                                }
                                else
                                {
                                    buildDescription.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecifications.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecifications.AlignmentPoint = centreOfBuildingSpecificationsOrto;
                                    buildDescription.AlignmentPoint = centreOfBuildDescriptionOrto;
                                }

                                currentSpace.AppendEntity(buildingSpecifications);
                                trans.AddNewlyCreatedDBObject(buildingSpecifications, true);
                                currentSpace.AppendEntity(buildDescription);
                                trans.AddNewlyCreatedDBObject(buildDescription, true);
                            }
                            else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "MTEXT" && buildDescriptionExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                            {
                                if (buildDescriptionWidth < segmentForRotateLength - 1.0)
                                {
                                    buildingSpecificationsMt.Location = centreOfBuildingSpecifications;
                                    buildDescriptionMt.Location = centreOfBuildDescription;
                                }
                                else
                                {
                                    buildDescriptionMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecificationsMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecificationsMt.Location = centreOfBuildingSpecificationsOrto;
                                    buildDescriptionMt.Location = centreOfBuildDescriptionOrto;
                                }
                                currentSpace.AppendEntity(buildingSpecificationsMt);
                                trans.AddNewlyCreatedDBObject(buildingSpecificationsMt, true);
                                currentSpace.AppendEntity(buildDescriptionMt);
                                trans.AddNewlyCreatedDBObject(buildDescriptionMt, true);
                            }
                            else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "TEXT" && buildDescriptionExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                            {

                                if (buildDescriptionWidth < segmentForRotateLength - 1.0)
                                {
                                    buildingSpecifications.AlignmentPoint = centreOfBuildingSpecifications;
                                    buildDescriptionMt.Location = centreOfBuildDescription;
                                }
                                else
                                {
                                    buildDescriptionMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecifications.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecifications.AlignmentPoint = centreOfBuildingSpecificationsOrto;
                                    buildDescriptionMt.Location = centreOfBuildDescriptionOrto;
                                }
                                currentSpace.AppendEntity(buildingSpecifications);
                                trans.AddNewlyCreatedDBObject(buildingSpecifications, true);
                                currentSpace.AppendEntity(buildDescriptionMt);
                                trans.AddNewlyCreatedDBObject(buildDescriptionMt, true);

                            }
                            else if (buildSpecificationExample.ObjectId.ObjectClass.DxfName == "MTEXT" && buildDescriptionExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                if (buildDescriptionWidth < segmentForRotateLength - 1.0)
                                {
                                    buildingSpecificationsMt.Location = centreOfBuildingSpecifications;
                                    buildDescription.AlignmentPoint = centreOfBuildDescription;
                                }
                                else
                                {
                                    buildDescription.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecificationsMt.Rotation = RotateAngleFix(angelOfRotation - (Math.PI / 2));
                                    buildingSpecificationsMt.Location = centreOfBuildingSpecificationsOrto;
                                    buildDescription.AlignmentPoint = centreOfBuildDescriptionOrto;
                                }
                                currentSpace.AppendEntity(buildingSpecificationsMt);
                                trans.AddNewlyCreatedDBObject(buildingSpecificationsMt, true);
                                currentSpace.AppendEntity(buildDescription);
                                trans.AddNewlyCreatedDBObject(buildDescription, true);
                            }
                        }
                        PromptResult buildNumb = doc.Editor.GetString("\nВведите номер дома, если требуется");
                        double numTextWeight = 0;
                        if (buildNumb.StringResult == "")
                        {
                            trans.Commit();
                            continue;
                        }
                        else
                        {
                            if (buildNumberExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                buildNumberText = buildNumberDBObject as DBText;
                                buildingNumber.TextString = buildNumb.StringResult;
                                buildingNumber.Color = buildNumberText.Color;
                                buildingNumber.Height = buildNumberText.Height;
                                buildingSpecificationsHeight = buildNumberText.Height;
                                buildingNumber.Layer = buildNumberText.Layer;
                                buildingNumber.TextStyleId = buildNumberText.TextStyleId;
                                Extents3d numbersBound = (Extents3d)buildingNumber.Bounds;
                                numTextWeight = numbersBound.MaxPoint[0] - numbersBound.MinPoint[0];
                                buildingNumber.Rotation = angelOfRotation;
                                buildingNumber.HorizontalMode = TextHorizontalMode.TextMid; //если использовать что-то кроме mid или centre текст создается в точке с нулевыми коорлдинатами

                            }
                            else if (buildNumberExample.ObjectId.ObjectClass.DxfName == "MTEXT")
                            {
                                buildNumberMText = buildNumberDBObject as MText;
                                buildingNumberMt.Contents = buildNumb.StringResult;
                                buildingNumberMt.Color = buildNumberMText.Color;
                                buildingNumberMt.TextHeight = buildNumberMText.TextHeight;
                                buildingSpecificationsHeight = buildNumberMText.TextHeight;
                                buildingNumberMt.Layer = buildNumberMText.Layer;
                                buildingNumberMt.TextStyleId = buildNumberMText.TextStyleId;
                                Extents3d numbersBound = (Extents3d)buildingNumberMt.Bounds;
                                numTextWeight = numbersBound.MaxPoint[0] - numbersBound.MinPoint[0];
                                buildingNumberMt.Rotation = angelOfRotation;
                                buildingNumberMt.Attachment = AttachmentPoint.MiddleCenter;

                            }
                            // Теперь вычисляем координаты подписи номера дома:
                            // 1. Сначала опредееляем какая из вершин отрезка ближайшего к улице, ближе к концу улицы (должен же пользователь как-то регулировать к какому углу дома будет ближе надпись):
                            Point3d pointForNumber = new Point3d();
                            Point3d anotherPoint = new Point3d();
                            Point3d segmentForRotateStart = new Point3d(segmentForRotate.StartPoint.X, segmentForRotate.StartPoint.Y, 0);
                            Point3d segmentForRotateEnd = new Point3d(segmentForRotate.EndPoint.X, segmentForRotate.EndPoint.Y, 0);
                            Point3d centreOfBuildingNumber = new Point3d();

                            //if (segmentForRotateStart.DistanceTo(streetEnd) < segmentForRotateEnd.DistanceTo(streetEnd))
                            if (segmentForRotateStart.DistanceTo(street.EndPoint) < segmentForRotateEnd.DistanceTo(street.EndPoint))
                            {
                                pointForNumber = segmentForRotateStart;
                                anotherPoint = segmentForRotateEnd;
                            }
                            else
                            {
                                pointForNumber = segmentForRotateEnd;
                                anotherPoint = segmentForRotateStart;
                            }

                            // 2. Зададём векторы для смещения относительно вершины
                            double numDy = buildingSpecificationsHeight * 0.7;
                            double numDx = numTextWeight / 2 + 0.25;
                            Vector3d moveIntoBuild = (centreOfBuilding - fromBuildToStreet).GetNormal() * numDy;
                            Vector3d moveBetweenCorners = (anotherPoint - pointForNumber).GetNormal() * numDx;
                            centreOfBuildingNumber = pointForNumber + moveBetweenCorners;
                            centreOfBuildingNumber = centreOfBuildingNumber + moveIntoBuild;
                            buildingNumber.AlignmentPoint = centreOfBuildingNumber;
                            currentSpace.AppendEntity(buildingNumber);
                            trans.AddNewlyCreatedDBObject(buildingNumber, true);
                        }
                        trans.Commit();
                    }
                }
            }
            doc.Editor.WriteMessage("\nПрограмма завершила работу");
        }
    }
}