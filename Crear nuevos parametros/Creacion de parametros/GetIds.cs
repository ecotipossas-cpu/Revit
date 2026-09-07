using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace GetIds
{
    [Transaction(TransactionMode.Manual)]
    public class GetIds : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // Verificar que la vista activa sea un Schedule
            ViewSchedule schedule = doc.ActiveView as ViewSchedule;

            if (schedule == null)
            {
                TaskDialog.Show(
                    "GetIds",
                    "Debes ejecutar el comando estando dentro de un Schedule."
                );

                return Result.Cancelled;
            }

            // Obtener los elementos que pertenecen al Schedule
            List<Element> scheduleElements =
                new FilteredElementCollector(doc, schedule.Id)
                    .WhereElementIsNotElementType()
                    .ToElements()
                    .ToList();

            if (scheduleElements.Count == 0)
            {
                TaskDialog.Show(
                    "GetIds",
                    "No se encontraron elementos en el Schedule."
                );

                return Result.Succeeded;
            }

            int updated = 0;
            int notFound = 0;

            using (Transaction tx = new Transaction(doc, "GetIds - Actualizar ElementID"))
            {
                tx.Start();

                foreach (Element element in scheduleElements)
                {
                    // Buscar parámetro ElementID
                    Parameter parameter = element.LookupParameter("ElementID");

                    if (parameter == null)
                    {
                        notFound++;
                        continue;
                    }

                    // Verificar que sea editable
                    if (parameter.IsReadOnly)
                    {
                        notFound++;
                        continue;
                    }

                    string elementId = element.Id.Value.ToString();

                    // Escribir según el tipo de parámetro
                    if (parameter.StorageType == StorageType.String)
                    {
                        parameter.Set(elementId);
                        updated++;
                    }
                    else if (parameter.StorageType == StorageType.Integer)
                    {
                        parameter.Set(element.Id.Value);
                        updated++;
                    }
                    else
                    {
                        notFound++;
                    }
                }

                tx.Commit();
            }

            string result =
    "GetIds completado.\n\n" +
    $"Elementos procesados: {scheduleElements.Count}\n" +
    $"ElementID actualizados: {updated}";

            if (notFound > 0)
            {
                result +=
                    $"\n\nElementID no encontrados/no editables: {notFound}";
            }

            TaskDialog.Show(
                "GetIds",
                result
            );

            return Result.Succeeded;
        }
    }
}