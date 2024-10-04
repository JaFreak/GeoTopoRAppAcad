
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
using System.Net.Mail;
using static System.Net.Mime.MediaTypeNames;

namespace WorkWithComunications
{
    public class CommunicationCase
    {
        public static string caseTextValue = null;
        public static double caseDiameter = 0.5;
        public static PromptEntityResult caseLineExample = null;
        public static PromptEntityResult caseTextExample = null;
        public static Polyline casePolylineExample = new Polyline();


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

        public Polyline CaseAxle() // создаем полилинию из указанных пользователем точек
        {
            Polyline сaseAxlePline = new Polyline();
            while (true)
            {
                PromptPointResult giveMePoint = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.GetPoint("\nУкажите точку по оси футляра. Нажмите Esc если достаточно");
                if (giveMePoint.Status != PromptStatus.OK)
                    break;
                else
                {
                    Point3d caseAxlePoint3d = giveMePoint.Value;
                    Point2d caseAxlePoint2d = new Point2d(caseAxlePoint3d.X, caseAxlePoint3d.Y);
                    сaseAxlePline.AddVertexAt(сaseAxlePline.NumberOfVertices, caseAxlePoint2d, 0, 0, 0);
                }
            }

            return сaseAxlePline;
        }

        public void CalculationAngleAndPointFromSegment2d(LineSegment2d takelineSegment2DHere, double offsetFromSegment, double objectHeight,
            out double angleBySegment, out Point3d pointInTheMidleOfSegment)      //здесь из сегмента,расстояния от него (в обе стороны) и высоты текста,
                                                                                  //расчитывается угол поворота объекта вдоль сегмента и координата точки в его середине
        {
            Vector2d caseAxleVector2D = (takelineSegment2DHere.StartPoint.GetVectorTo(takelineSegment2DHere.MidPoint));
            Vector2d caseAxlePerpVector2D = caseAxleVector2D.GetPerpendicularVector().GetNormal();
            caseAxlePerpVector2D = caseAxlePerpVector2D * ((offsetFromSegment / 2) + objectHeight);
            Point2d pointForText2D = takelineSegment2DHere.StartPoint + caseAxleVector2D;
            pointForText2D = pointForText2D + caseAxlePerpVector2D;
            pointInTheMidleOfSegment = new Point3d(pointForText2D.X, pointForText2D.Y, 0);

            Point3d segmentStartPoint = new Point3d(takelineSegment2DHere.StartPoint.X, takelineSegment2DHere.StartPoint.Y, 0);
            Point3d segmentEndPoint = new Point3d(takelineSegment2DHere.EndPoint.X, takelineSegment2DHere.EndPoint.Y, 0);
            Vector3d caseAxleVector3D = segmentStartPoint.GetVectorTo(segmentEndPoint);
            double angelOfRotation = caseAxleVector3D.GetAngleTo(Vector3d.XAxis, -Vector3d.ZAxis);
            angleBySegment = RotateAngleFix(angelOfRotation);
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


        [CommandMethod("CaseCreate")]
        [CommandMethod("СоздатьФутляр")]
        public void ComunicationCaseCreate()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database dB = doc.Database;
            Editor ed = doc.Editor;
            // Starts a new transaction with the Transaction Manager
            using (Transaction trans = dB.TransactionManager.StartTransaction())
            {
                // Получение доступа к активному пространству (пространство модели или лист)
                BlockTableRecord currentSpace = trans.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                // Предлагаем пользователю выбрать примеры примитивов подписей
                if (caseLineExample == null)
                {
                    caseLineExample = ed.GetEntity("\nВыберите полилинию - пример контура футляра.");

                    if (caseLineExample.ObjectId == ObjectId.Null)
                    {
                        doc.Editor.WriteMessage("\nНичего не выбрано. Программа прекратила работу");
                        return;
                    }
                    else if (caseLineExample.ObjectId.ObjectClass.DxfName != "POLYLINE" & caseLineExample.ObjectId.ObjectClass.DxfName != "LWPOLYLINE")
                    {
                        while (caseLineExample.ObjectId.ObjectClass.DxfName != "POLYLINE" & caseLineExample.ObjectId.ObjectClass.DxfName != "LWPOLYLINE")
                            caseLineExample = ed.GetEntity("\nВыберите полилинию - пример контура футляра.");
                    }
                    casePolylineExample = caseLineExample.ObjectId.GetObject(OpenMode.ForWrite) as Polyline;
                }

                if (caseTextExample == null)
                {
                    caseTextExample = ed.GetEntity("\nВыберите пример подписи футляра. Текст или Мтекст");
                    if (caseTextExample.ObjectId == ObjectId.Null)
                    {
                        doc.Editor.WriteMessage("\nНичего не выбрано. Программа прекратила работу");
                        return;
                    }
                    else if (caseTextExample.ObjectId.ObjectClass.DxfName != "TEXT" & caseTextExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                    {
                        while (caseTextExample.ObjectId.ObjectClass.DxfName != "TEXT" & caseTextExample.ObjectId.ObjectClass.DxfName != "MTEXT")
                            caseTextExample = ed.GetEntity("\nВыберите пример подписи футляра. Текст или Мтекст");
                    }
                }
                trans.Commit();
            }

            // Starts a new transaction with the Transaction Manager
            using (Transaction trans2 = dB.TransactionManager.StartTransaction())
            {
                DBText caseSpecExT = new DBText();
                MText caseSpecExMt = new MText();
                double specificationHeight = 0;

                if (caseTextExample.ObjectId.ObjectClass.DxfName == "TEXT")
                {
                    caseSpecExT = caseTextExample.ObjectId.GetObject(OpenMode.ForWrite) as DBText;
                    specificationHeight = caseSpecExT.Height;
                }
                else
                {
                    caseSpecExMt = caseTextExample.ObjectId.GetObject(OpenMode.ForWrite) as MText;
                    specificationHeight = caseSpecExMt.TextHeight;
                }

                BlockTableRecord currentSpace2 = trans2.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                DBObjectCollection caseCreateCollection = new DBObjectCollection();
                Polyline caseAxle = CaseAxle();
                if (caseAxle == null)
                    return;
                else
                {


                    if (caseAxle.NumberOfVertices < 2)
                    {
                        ed.WriteMessage("\nНеобходимо больше 1 точки. Программа прекратила работу");
                        return;
                    }
                    else
                    {
                        Polyline casePline = new Polyline();
                        Polyline casePline1 = new Polyline();
                        Polyline casePline2 = new Polyline();
                        PromptDoubleOptions caseDiameterOptions = new PromptDoubleOptions("\nВведите диаметр футляра")
                        {
                            DefaultValue = caseDiameter,
                            UseDefaultValue = true
                        };
                        PromptDoubleResult caseDiameterValue = ed.GetDouble(caseDiameterOptions);
                        if (caseDiameterValue == null)
                            caseDiameter = 0.5;
                        else
                            caseDiameter = caseDiameterValue.Value;

                        caseCreateCollection = caseAxle.GetOffsetCurves(caseDiameter / 2);
                        casePline1 = (Polyline)caseCreateCollection[0];
                        caseCreateCollection = caseAxle.GetOffsetCurves(-caseDiameter / 2);
                        casePline2 = (Polyline)caseCreateCollection[0];
                        int z = casePline1.NumberOfVertices;
                        for (int i = 0; i < casePline1.NumberOfVertices; i++)
                        {
                            casePline.AddVertexAt(i, casePline1.GetPoint2dAt(i), 0, 0, 0);
                        }
                        for (int y = (casePline2.NumberOfVertices - 1); y >= 0; y--)
                        {
                            casePline.AddVertexAt(z, casePline2.GetPoint2dAt(y), 0, 0, 0);
                            z++;
                        }
                        casePline.Elevation = 0;
                        casePline.Color = casePolylineExample.Color;
                        casePline.LayerId = casePolylineExample.LayerId;
                        casePline.Linetype = casePolylineExample.Linetype;
                        casePline.LineWeight = casePolylineExample.LineWeight;
                        casePline.LinetypeScale = casePolylineExample.LinetypeScale;
                        casePline.Plinegen = false;
                        casePline.Closed = true;

                        currentSpace2.AppendEntity(casePline);
                        trans2.AddNewlyCreatedDBObject(casePline, true);

                        PromptStringOptions caseSpecOptions = new PromptStringOptions("\nВведите характеристики,если требуется.Ввод оканчивается нажатием клавиши Enter")
                        {
                            AllowSpaces = true,
                            DefaultValue = caseTextValue,
                            UseDefaultValue = true
                        };
                        PromptResult caseSpecValue = doc.Editor.GetString(caseSpecOptions);
                        if (caseSpecValue.StringResult == "")
                            caseTextValue = null;
                        else
                        {
                            caseTextValue = caseSpecValue.StringResult;

                            if (caseTextExample.ObjectId.ObjectClass.DxfName == "TEXT")
                            {
                                for (int s = 0; s < caseAxle.NumberOfVertices - 1; s++)
                                {
                                    DBText[] caseSpecifications = new DBText[caseAxle.NumberOfVertices - 1];
                                    DBText caseSpecEx = caseTextExample.ObjectId.GetObject(OpenMode.ForWrite) as DBText;
                                    caseSpecifications[s] = new DBText();
                                    caseSpecifications[s].TextString = caseTextValue;
                                    caseSpecifications[s].HorizontalMode = TextHorizontalMode.TextMid;
                                    caseSpecifications[s].Color = caseSpecEx.Color;
                                    caseSpecifications[s].Height = caseSpecEx.Height;
                                    caseSpecifications[s].Layer = caseSpecEx.Layer;
                                    caseSpecifications[s].TextStyleId = caseSpecEx.TextStyleId;
                                    caseSpecifications[s].LineWeight = caseSpecEx.LineWeight;
                                    caseSpecifications[s].Linetype = caseSpecEx.Linetype;

                                    LineSegment2d casePart = caseAxle.GetLineSegment2dAt(s);
                                    CalculationAngleAndPointFromSegment2d(casePart, caseDiameter, specificationHeight, out double rotatAngle, out Point3d pointForText);
                                    caseSpecifications[s].Rotation = rotatAngle;
                                    caseSpecifications[s].AlignmentPoint = pointForText;
                                    currentSpace2.AppendEntity(caseSpecifications[s]);
                                    trans2.AddNewlyCreatedDBObject(caseSpecifications[s], true);
                                }
                            }
                            else
                            {
                                for (int s = 0; s < caseAxle.NumberOfVertices - 1; s++)
                                {
                                    MText[] caseSpecifications = new MText[caseAxle.NumberOfVertices - 1];
                                    MText caseSpecEx = caseTextExample.ObjectId.GetObject(OpenMode.ForWrite) as MText;
                                    caseSpecifications[s] = new MText();
                                    caseSpecifications[s].Contents = caseTextValue;
                                    caseSpecifications[s].Attachment = AttachmentPoint.MiddleCenter;
                                    caseSpecifications[s].Color = caseSpecEx.Color;
                                    caseSpecifications[s].Height = caseSpecEx.Height;
                                    caseSpecifications[s].Layer = caseSpecEx.Layer;
                                    caseSpecifications[s].TextStyleId = caseSpecEx.TextStyleId;
                                    caseSpecifications[s].TextHeight = caseSpecEx.TextHeight;
                                    caseSpecifications[s].BackgroundFill = caseSpecEx.BackgroundFill;
                                    caseSpecifications[s].BackgroundFillColor = caseSpecEx.BackgroundFillColor;
                                    caseSpecifications[s].BackgroundScaleFactor = caseSpecEx.BackgroundScaleFactor;
                                    caseSpecifications[s].LineWeight = caseSpecEx.LineWeight;
                                    caseSpecifications[s].Linetype = caseSpecEx.Linetype;

                                    LineSegment2d casePart = caseAxle.GetLineSegment2dAt(s);
                                    CalculationAngleAndPointFromSegment2d(casePart, caseDiameter, specificationHeight, out double rotatAngle, out Point3d pointForText);
                                    caseSpecifications[s].Location = pointForText;
                                    caseSpecifications[s].Rotation = rotatAngle;
                                    currentSpace2.AppendEntity(caseSpecifications[s]);
                                    trans2.AddNewlyCreatedDBObject(caseSpecifications[s], true);
                                }
                            }
                        }
                    }

                }
                trans2.Commit();
            }

            ed.WriteMessage("\nПрограмма завершила работу");
        }
    }
}