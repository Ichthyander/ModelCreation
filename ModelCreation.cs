using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.ApplicationServices;

namespace ModelCreation
{
    [TransactionAttribute(TransactionMode.Manual)]
    class ModelCreation : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document document = commandData.Application.ActiveUIDocument.Document;

            var levelsList = GetLevels(document);

            var baseLevel = levelsList
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();

            var level2 = levelsList
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault();

            List<Wall> walls = CreateWalls(document, baseLevel, level2);

            AddDoor(document, baseLevel, walls[0]);
            AddWindows(document, baseLevel, walls.GetRange(1, 3));
            AddRoof(document, level2, walls);

            return Result.Succeeded;
        }

        private void AddRoof(Document document, Level roofLevel, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(document)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Типовой - 400мм"))
                .Where(x => x.FamilyName.Equals("Базовая крыша"))
                .FirstOrDefault();

            double wallWidth = walls[0].Width;
            double dt = wallWidth / 2;

            //List<XYZ> points = new List<XYZ>();
            //points.Add(new XYZ(-dt, -dt, 0));
            //points.Add(new XYZ(dt, -dt, 0));
            //points.Add(new XYZ(dt, dt, 0));
            //points.Add(new XYZ(-dt, dt, 0));
            //points.Add(new XYZ(-dt, -dt, 0));

            Application application = document.Application;
            CurveArray curveArray = application.Create.NewCurveArray();
            LocationCurve curve = walls[0].Location as LocationCurve;
            XYZ p1 = new XYZ(curve.Curve.GetEndPoint(0).X, curve.Curve.GetEndPoint(0).Y, walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble());
            
            XYZ p2 = new XYZ((curve.Curve.GetEndPoint(0).X + curve.Curve.GetEndPoint(1).X)/2.0, 
                (curve.Curve.GetEndPoint(0).Y + curve.Curve.GetEndPoint(1).Y) / 2.0, 
                walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble() + UnitUtils.ConvertToInternalUnits(1500, DisplayUnitType.DUT_MILLIMETERS));

            XYZ p3 = new XYZ(curve.Curve.GetEndPoint(1).X, curve.Curve.GetEndPoint(1).Y, walls[0].get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble());

            curveArray.Append(Line.CreateBound(p1, p2));
            curveArray.Append(Line.CreateBound(p2, p3));

            View view = new FilteredElementCollector(document)
                .OfClass(typeof(View))
                .Cast<View>()
                .Where(x => x.Name.Equals("Южный"))
                .FirstOrDefault();


            Transaction transaction = new Transaction(document, "Create ExtrusionRoof");
            transaction.Start();

            ReferencePlane plane = document.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(1, 0, 0), new XYZ(0, 0, 1), view);

            document.Create.NewExtrusionRoof(curveArray, plane, roofLevel, roofType, 
                -1*(walls[1].Location as LocationCurve).Curve.ApproximateLength / 2 - dt, (walls[1].Location as LocationCurve).Curve.ApproximateLength / 2 + dt);

            transaction.Commit();
        }

        //private void AddRoof(Document document, Level roofLevel, List<Wall> walls)
        //{
        //    RoofType roofType = new FilteredElementCollector(document)
        //        .OfClass(typeof(RoofType))
        //        .OfType<RoofType>()
        //        .Where(x => x.Name.Equals("Типовой - 400мм"))
        //        .Where(x => x.FamilyName.Equals("Базовая крыша"))
        //        .FirstOrDefault();

        //    double wallWidth = walls[0].Width;
        //    double dt = wallWidth / 2;

        //    List<XYZ> points = new List<XYZ>();
        //    points.Add(new XYZ(-dt, -dt, 0));
        //    points.Add(new XYZ(dt, -dt, 0));
        //    points.Add(new XYZ(dt, dt, 0));
        //    points.Add(new XYZ(-dt, dt, 0));
        //    points.Add(new XYZ(-dt, -dt, 0));

        //    Application application = document.Application;
        //    CurveArray footprint = application.Create.NewCurveArray();

        //    for (int i = 0; i < 4; i++)
        //    {
        //        LocationCurve curve = walls[i].Location as LocationCurve;
        //        XYZ p1 = curve.Curve.GetEndPoint(0);
        //        XYZ p2 = curve.Curve.GetEndPoint(1);
        //        Line line = Line.CreateBound(p1 + points[i], p2 + points[i + 1]);
        //        footprint.Append(line);
        //    }

        //    Transaction transaction = new Transaction(document, "Create Roof");
        //    transaction.Start();
        //    ModelCurveArray footprintToModelCurveMapping = new ModelCurveArray();
        //    FootPrintRoof footprintRoof = document.Create.NewFootPrintRoof(footprint, roofLevel, roofType, out footprintToModelCurveMapping);

        //    foreach(ModelCurve modelCurve in footprintToModelCurveMapping)
        //    {
        //        footprintRoof.set_DefinesSlope(modelCurve, true);
        //        footprintRoof.set_SlopeAngle(modelCurve, 0.5);
        //    }

        //    transaction.Commit();
        //}

        private void AddWindows(Document document, Level baseLevel, List<Wall> walls)
        {
            FamilySymbol windowType = new FilteredElementCollector(document)
                  .OfClass(typeof(FamilySymbol))
                  .OfCategory(BuiltInCategory.OST_Windows)
                  .OfType<FamilySymbol>()
                  .Where(x => x.Name.Equals("0915 x 1220 мм"))
                  .Where(x => x.FamilyName.Equals("Фиксированные"))
                  .FirstOrDefault();

            Transaction transaction = new Transaction(document, "Add window");
            transaction.Start();

            if (!windowType.IsActive)
                windowType.Activate();

            foreach (var wall in walls)
            {
                LocationCurve hostCurve = wall.Location as LocationCurve;
                XYZ point1 = hostCurve.Curve.GetEndPoint(0);
                XYZ point2 = hostCurve.Curve.GetEndPoint(1);
                XYZ middlePoint = (point1 + point2) / 2;
                XYZ placementPoint = new XYZ(middlePoint.X, middlePoint.Y, wall.get_Parameter(BuiltInParameter.WALL_USER_HEIGHT_PARAM).AsDouble() / 2);

                document.Create.NewFamilyInstance(placementPoint, windowType, wall, baseLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            }

            transaction.Commit();
        }

        private void AddDoor(Document document, Level baseLevel, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(document)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ placementPoint = (point1 + point2) / 2;

            Transaction transaction = new Transaction(document, "Add door");
            transaction.Start();

            if (!doorType.IsActive)
                doorType.Activate();

            document.Create.NewFamilyInstance(placementPoint, doorType, wall, baseLevel, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            transaction.Commit();
        }

        public List<Level> GetLevels(Document document)
        {
            return new FilteredElementCollector(document)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();
        }

        public List<Wall> CreateWalls(Document document, Level level1, Level level2)
        {
            double width = UnitUtils.ConvertToInternalUnits(10000, DisplayUnitType.DUT_MILLIMETERS);
            double depth = UnitUtils.ConvertToInternalUnits(5000, DisplayUnitType.DUT_MILLIMETERS);

            double dx = width / 2;
            double dy = depth / 2;

            List<XYZ> points = new List<XYZ>();
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));

            List<Wall> walls = new List<Wall>();

            Transaction transaction = new Transaction(document, "Create walls");
            transaction.Start();

            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(document, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }

            transaction.Commit();

            return walls;
        }
    }
}
