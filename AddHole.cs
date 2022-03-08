using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddHole
{
    [Transaction(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        private Document arDoc;
        private Document ovDoc;
        private View3D view3d;
        private FamilySymbol holeSymbol;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDocument = uiApp.ActiveUIDocument;
            arDoc = uiDocument.Document;

            ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                message = "Не найден новый файл";
                return Result.Failed;
            }


            holeSymbol = new FilteredElementCollector(arDoc)
               .OfClass(typeof(FamilySymbol))
               .OfCategory(BuiltInCategory.OST_GenericModel)
               .OfType<FamilySymbol>()
               .Where(x => x.FamilyName.Equals("Отверстие"))
               .FirstOrDefault();

            if (holeSymbol == null)
            {
                message = "Не найдено семейство";
                return Result.Failed;
            }



            view3d = new FilteredElementCollector(arDoc)
               .OfClass(typeof(View3D))
               .OfType<View3D>()
               .Where(x => !x.IsTemplate)
               .FirstOrDefault();

            if (view3d == null)
            {
                message = "Не найден view3d";
                return Result.Failed;
            }



            AddHolesForT<Duct>();
            AddHolesForT<Pipe>();
            return Result.Succeeded;
        }

        private void AddHolesForT<T>()
        {
            //T is Duct or Pipe


            List<MEPCurve> mepObjects = new FilteredElementCollector(ovDoc)
                 .OfClass(typeof(T))
                 .OfCategory(typeof(T).Equals(typeof(Duct)) ? BuiltInCategory.OST_DuctCurves : BuiltInCategory.OST_PipeCurves)
                 .OfType<MEPCurve>()
                 .ToList();


            ReferenceIntersector referenceIntersector = new ReferenceIntersector(
        new ElementClassFilter(typeof(Wall)),
         FindReferenceTarget.Element, view3d);

            using (var ts = new Transaction(arDoc, $"adding holes for {typeof(T).FullName}"))
            {
                ts.Start();
                foreach (MEPCurve mepObj in mepObjects)
                {
                    Curve curve = (mepObj.Location as LocationCurve).Curve;
                    XYZ point = curve.GetEndPoint(0);
                    XYZ dir = (curve as Line).Direction;

                    List<ReferenceWithContext> intersections = referenceIntersector.Find(point, dir)
                        .Where(x => x.Proximity <= curve.Length)
                        .Distinct(new ReferenceWithContextElementEqualityComparer())
                        .ToList();

                    foreach (ReferenceWithContext context in intersections)
                    {
                        double proximity = context.Proximity;
                        Reference reference = context.GetReference();
                        Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                        Level level = arDoc.GetElement(wall.LevelId) as Level;
                        XYZ pointHole = point + (dir * proximity);

                        FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, holeSymbol, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                        
                        Parameter width = hole.LookupParameter("Ширина");
                        Parameter height = hole.LookupParameter("Высота");
                        double offset = UnitUtils.ConvertToInternalUnits(50, UnitTypeId.Millimeters);
                        if (mepObj.get_Parameter(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM) != null || mepObj.get_Parameter(BuiltInParameter.RBS_PIPE_DIAMETER_PARAM) != null)
                        {
                            width?.Set(mepObj.Diameter + offset);
                            height?.Set(mepObj.Diameter + offset);
                        }
                        else
                        {
                            width?.Set(mepObj.get_Parameter(BuiltInParameter.RBS_CURVE_WIDTH_PARAM).AsDouble() + offset);
                            height?.Set(mepObj.get_Parameter(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM).AsDouble() + offset);
                        }


                    }
                }
                ts.Commit();
            }
        }

        private void debug(Type f, string v)
        {
            TaskDialog.Show("T=" + f.Name, v);
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

}
