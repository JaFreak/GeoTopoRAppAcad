
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
using RMMethods;


namespace WorkWithComunications
{
    public class CommunicationCase
    {
        public static string caseTextValue = null;
        public static double caseDiameter = 0.5;
        public static PromptEntityResult caseLineExample = null;
        public static PromptEntityResult caseTextExample = null;
        public static Polyline casePolylineExample = new Polyline();

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
                Polyline caseAxle = RM.AxlePolyline();
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
                                    RM.CalculationAngleAndPointFromSegment2d(casePart, caseDiameter, specificationHeight, out double rotatAngle, out Point3d pointForText);
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
                                    RM.CalculationAngleAndPointFromSegment2d(casePart, caseDiameter, specificationHeight, out double rotatAngle, out Point3d pointForText);
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