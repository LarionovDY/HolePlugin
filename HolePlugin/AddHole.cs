using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HolePlugin
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //обращение к основному документу (АР)
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            //обращение к связанному файлу (ОВ) через поиск по названию файла
            Document ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }
            //проверка что семейство проёмов загружено в модель АР
            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстия"))
                .FirstOrDefault();            
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }
            //поиск 3D вида
            View3D view3D = new FilteredElementCollector(arDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(x => !x.IsTemplate) //проверка что найденный вид не является шаблоном
                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            //поиск воздуховодов в модели
            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();
            
            //поиск пересечений стен воздуховодами
            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);

            Transaction transaction0 = new Transaction(arDoc);
            transaction0.Start("Активация FamilySymbol");
            //активация FamilySymbol
            if (!familySymbol.IsActive)
                familySymbol.Activate();
            transaction0.Commit();  

            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Расстановка отверстий");           
            foreach (Duct d in ducts)
            {
                //получение кривой проекции воздуховода
                Line curve = (d.Location as LocationCurve).Curve as Line;
                //получение точки начала кривой
                XYZ point = curve.GetEndPoint(0);
                //получение вектора направления кривой
                XYZ direction = curve.Direction;
                //метод поиска пересечений через точку и направление
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //среди всех объектов которые совпадают по критерию осталяют один
                    .ToList();
                foreach (ReferenceWithContext refer in intersections)
                {
                    //расстояние до пересечения
                    double proximity = refer.Proximity;
                    //ссылка на элемент
                    Reference reference = refer.GetReference();
                    //получение из ссылки элемента (стены)
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    //получение уровня на котором находится стена
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    //получение точки вставки отверстия
                    XYZ pointHole = point + (direction * proximity); 
                    //вставка проёмов в стены
                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    //задание размеров проёма
                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(d.Diameter);
                    height.Set(d.Diameter);
                }
            }
            transaction.Commit();
            return Result.Succeeded;
        }
    }
    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        //проверка будут ли два заданных объекта одинаковыми
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
        //возвращает хэш-код объекта
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
