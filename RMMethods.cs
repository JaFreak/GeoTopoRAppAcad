
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

namespace RMMethods
{
    public class RM
    {
        //не допускает значение угла, при назначении которого объект будет показан на плане вверх ногами
        public static double RotateAngleFix(double angle)      
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
        public static Polyline3d Axle3dPolyline()
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
        // создает полилинию из указанных пользователем точек
        public static Polyline AxlePolyline() 
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

        public static void CalculationAngleAndPointFromSegment2d(LineSegment2d takelineSegment2DHere, double offsetFromSegment, double objectHeight,
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
        public static bool EntitysBoundIntersectCheck(Point3d centreOfEntity, Entity obj1, Entity obj2)
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

    }
}