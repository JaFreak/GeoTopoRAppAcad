
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
        public Polyline CaseAxle()
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
                // Получение доступа к активному пространству (пространство модели или лист)
                BlockTableRecord currentSpace2 = trans2.GetObject(doc.Database.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;
                DBObjectCollection caseCreateCollection = new DBObjectCollection();
                Polyline caseAxle = CaseAxle();
                if (caseAxle == null)
                    return;
                else
                {
                    Vector3d caseAxleVector3D = (caseAxle.GetPoint3dAt(0).GetVectorTo(caseAxle.GetPoint3dAt(1)));
                    caseAxleVector3D = caseAxleVector3D.GetNormal() * caseAxle.GetPoint3dAt(0).DistanceTo(caseAxle.GetPoint3dAt(1));
                    Vector3d caseAxlePerpVector3D = caseAxleVector3D.GetPerpendicularVector().GetNormal();
                    double angelOfRotation = caseAxleVector3D.GetAngleTo(Vector3d.XAxis, -Vector3d.ZAxis);
                    double rotatAngle = RotateAngleFix(angelOfRotation);

                    if (caseAxle.NumberOfVertices < 2)
                    {
                        ed.WriteMessage("\nНеобходимо больше 1 точки. Программа прекратила работу");
                        return;
                    }
                    else
                    {
                        //открываем для чтения

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

                        caseAxlePerpVector3D = caseAxlePerpVector3D * ((caseDiameter / 2) + 0.3);
                        Point3d pointForText = caseAxle.GetPoint3dAt(0) + caseAxleVector3D;
                        pointForText = pointForText + caseAxlePerpVector3D;

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
                                DBText caseSpecEx = caseTextExample.ObjectId.GetObject(OpenMode.ForWrite) as DBText;
                                DBText caseSpec = new DBText();
                                caseSpec.TextString = caseTextValue;
                                caseSpec.HorizontalMode = TextHorizontalMode.TextMid;
                                caseSpec.Rotation = rotatAngle;
                                caseSpec.AlignmentPoint = pointForText;
                                caseSpec.Color = caseSpecEx.Color;
                                caseSpec.Height = caseSpecEx.Height;
                                caseSpec.Layer = caseSpecEx.Layer;
                                caseSpec.TextStyleId = caseSpecEx.TextStyleId;
                                currentSpace2.AppendEntity(caseSpec);
                                trans2.AddNewlyCreatedDBObject(caseSpec, true);
                            }
                            else
                            {
                                MText caseSpecEx = caseTextExample.ObjectId.GetObject(OpenMode.ForWrite) as MText;
                                MText caseSpec = new MText();
                                caseSpec.Contents = caseTextValue;
                                caseSpec.Attachment = AttachmentPoint.MiddleCenter;
                                caseSpec.Location = pointForText;
                                caseSpec.Rotation = rotatAngle;
                                caseSpec.Color = caseSpecEx.Color;
                                caseSpec.Height = caseSpecEx.Height;
                                caseSpec.Layer = caseSpecEx.Layer;
                                caseSpec.TextStyleId = caseSpecEx.TextStyleId;
                                currentSpace2.AppendEntity(caseSpec);
                                trans2.AddNewlyCreatedDBObject(caseSpec, true);
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