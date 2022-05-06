using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreationModelPlugin_3
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModelPlugin_3 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            Level level1, level2;

            GetLevels(doc, out level1, out level2);
            CreateBuilding(doc, level1, level2);

            return Result.Succeeded;
        }

        private static void GetLevels(Document doc, out Level level1, out Level level2)
        {
            List<Level> listLevel = new FilteredElementCollector(doc) //фильтруем все уровни документа в список
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();

            level1 = listLevel
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();  //берем только 1 уровень из коллекции
            level2 = listLevel
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault(); //берем только 2 уровень из коллекции
        }

        private static void CreateBuilding(Document doc, Level level1, Level level2)
        {
            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters); //зададим ширину дома, приведя ее от 10000мм к системным еденицам
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters); //зададим глубину дома, приведя ее от 5000мм к системным еденицам
            double dx = width / 2;
            double dy = depth / 2; //получать координаты мы решили от центра нашего дома

            List<XYZ> points = new List<XYZ>(); //создаем список координат (помним, что координаты от центра дома)
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0)); //добавили пятую точку ровно аналогичную первой для простоты, т.к. строить стены мы будем в цикле, перебирая точки попарно: 1-2, 2-3, 3-4, 4-5

            List<Wall> walls = new List<Wall>(); //создали массив под стены на будущее

            Transaction tr = new Transaction(doc); //запускаем транзакцию для добавления в модель
            tr.Start("Построение здания");
            for (int i = 0; i < 4; i++) //перебираем циклом точки
            {
                Line line = Line.CreateBound(points[i], points[i + 1]); //на основании точек создаем линию
                Wall wall = Wall.Create(doc, line, level1.Id, false); // по линии создаем стену. Использовали .Id
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id); //задаем высоту для каждой стены через отдельный параметр
                walls.Add(wall); //добавляем стену в массив на будущее                
            }

            AddDoor(doc, level1, walls[0]); //добавляем дверь методом в первую стену из списка

            for (int i = 0; i < 3; i++)
            {
                AddWindow(doc, level1, walls[i + 1]);
            }

            tr.Commit();
        }

        private static void AddDoor(Document doc, Level level1, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc) //отфильтровываем из документа необходимое семейство и типоразмер двери
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve; // Location для стены - это кривая Curve, поэтому преобразуем ее к этому типу
            XYZ startPoint = hostCurve.Curve.GetEndPoint(0); //находим точку начала кривой, где 0 - самая первая точка
            XYZ endPoint = hostCurve.Curve.GetEndPoint(1); //находим точку конца кривой
            XYZ centerPoint = (startPoint + endPoint) / 2; //находим центр кривой, совпадает с центром проекции стены

            if (!doorType.IsActive)
                doorType.Activate(); //такую запись мы делаем по подсказку Jeremy Tammick т.к. необходимо проверить активен ли такой тип втсавляемго элемента
                                     //в документе и если нет, то активировать

            doc.Create.NewFamilyInstance(centerPoint, doorType, wall, level1, StructuralType.NonStructural); //создаем экземпляр двери в модели
        }

        private static void AddWindow(Document doc, Level level1, Wall wall)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc) //отфильтровываем из документа необходимое семейство и типоразмер двери
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0610 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve; //'это бы тоже в отдельный метод вынести
            XYZ startPoint = hostCurve.Curve.GetEndPoint(0);
            XYZ endPoint = hostCurve.Curve.GetEndPoint(1);
            XYZ centerPoint = (startPoint + endPoint) / 2;

            if (!windowType.IsActive)
                windowType.Activate();

            Element window = doc.Create.NewFamilyInstance(centerPoint, windowType, wall, level1, StructuralType.NonStructural); //создаем экземпляр окна в модели
            Parameter height = window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
            double heightMM = UnitUtils.ConvertToInternalUnits(850, UnitTypeId.Millimeters); //где 850 - высота от пола до низа окна
            height.Set(heightMM); //поднимаем окна от пола
        }

        private static void AddRoof(Document doc, Level level2, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(doc) //выбираем тип крыши, которую будем создавать
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(x => x.Name.Equals("Типовой - 400мм"))
                .Where(x => x.FamilyName.Equals("Базовая крыша"))
                .FirstOrDefault();

            CurveArray curveArray = new CurveArray();
            curveArray.Append(Line.CreateBound(new XYZ(-20, -10, 13), new XYZ(-20, 0, 20)));
            curveArray.Append(Line.CreateBound(new XYZ(-20, 0, 20), new XYZ(-20, 10, 13)));

            ReferencePlane plane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 20), new XYZ(0, 20, 0), doc.ActiveView);
            doc.Create.NewExtrusionRoof(curveArray, plane, level2, roofType, -19, 19);

        }
    }
}
