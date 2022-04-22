using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin_8
{[Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document; // 1-й документ,АР
            Document ovDoc = arDoc.Application.Documents
                .OfType<Document>() //фильтрация док-а
                .Where(x => x.Title.Contains("ОВ"))
                .FirstOrDefault(); // связанный файл ОВ
            if (ovDoc == null) // если файл ОВ не обнаружен 
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;  //закроем плагин с отменой
            }

           
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)   //отверстия
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)   //поиск по категории
                .OfType<FamilySymbol>()   //приведём к типу
                .Where(x => x.FamilyName.Equals("Отверстие"))   //поиск по названию
                .FirstOrDefault();
            if (familySymbol == null)  //проверка
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");   //сообщим
                return Result.Cancelled; //выйдем
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc) // найдем все воздуховоды, вернём лист
                .OfClass(typeof(Duct)) //фильтр по классу
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc) // найдем все трубы
               .OfClass(typeof(Pipe))
               .OfType<Pipe>()
               .ToList();

            View3D view3D = new FilteredElementCollector(arDoc) //  3Д вид
                .OfClass(typeof(View3D))   //фильтр по классу
                .OfType<View3D>()
                .Where(x => !x.IsTemplate)  //свойство из .мсд не установлено
                .FirstOrDefault();

            if (view3D == null)  //проверка
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");   //сообщить
                return Result.Cancelled;  //выйдем
            }

          
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)),
                FindReferenceTarget.Element, view3D);  //создание объекта

            Transaction transaction = new Transaction(arDoc); //транзакция
            transaction.Start("Расстановка отверстий");
            foreach (Duct duct in ducts)
            {
                Line curve = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);// исходная точка
                XYZ direction = curve.Direction;// векор направления

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) //набор всех пересечений
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())  //метод расширения
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)   //переберём все пересечения
                {
                    double proximity = refer.Proximity; // расстояние которые не превышают длину воздуховодов
                    Reference reference = refer.GetReference(); // ссылка внутренняя по Id
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall; //стена на которую идет ссылка
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //из уровня выводим объект
                    XYZ pointHole = point + (direction * proximity); //стартовая точка, направление+растояние

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, 
                        familySymbol, wall, level, 
                        StructuralType.NonStructural); // вставка объекта
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(duct.Diameter);  //значение, которое возьмём у воздуховода
                    height.Set(duct.Diameter);
                }
            }


            foreach (Pipe pipe in pipes)
            {
                Line curve = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(pipe.Diameter);
                    height.Set(pipe.Diameter);
                }
            }
            transaction.Commit();
            return Result.Succeeded;
        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                       && xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
    
}
