using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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

            return Result.Succeeded;
        }

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
