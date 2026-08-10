using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;

namespace BimDataSync
{
    [Transaction(TransactionMode.Manual)]
    public class BimDataSyncCommand : IExternalCommand
    {
        // ============================================================
        // RUTA DEL ARCHIVO EXCEL
        // ============================================================

        private const string ExcelPath =
            @"E:\PC\Escritorio\CORREDORES VIALES CUNDINAMARCA\C#\Asignar data a elementos\Data contenciones.xlsx";


        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --------------------------------------------------------
            // Verificar archivo
            // --------------------------------------------------------

            if (!File.Exists(ExcelPath))
            {
                TaskDialog.Show(
                    "BimDataSync",
                    "No se encontró el archivo Excel:\n\n" + ExcelPath);

                return Result.Failed;
            }

            int hojasProcesadas = 0;
            int filasProcesadas = 0;
            int elementosEncontrados = 0;
            int valoresAsignados = 0;
            int parametrosNoEncontrados = 0;
            int elementosNoEncontrados = 0;
            int parametrosSoloLectura = 0;
            int errores = 0;

            List<string> logErrores = new List<string>();

            // NUEVO: lista detallada de parámetros no encontrados
            List<ParametroNoEncontrado> parametrosNoEncontradosLista =
                new List<ParametroNoEncontrado>();

            // --------------------------------------------------------
            // Abrir Excel
            // --------------------------------------------------------

            using (XLWorkbook workbook = new XLWorkbook(ExcelPath))
            {
                // ----------------------------------------------------
                // TRANSACTION
                // ----------------------------------------------------

                using (Transaction trans = new Transaction(
                    doc,
                    "BimDataSync - Asignar datos"))
                {
                    trans.Start();

                    // =================================================
                    // RECORRER TODAS LAS HOJAS
                    // =================================================

                    foreach (IXLWorksheet worksheet in workbook.Worksheets)
                    {
                        hojasProcesadas++;

                        IXLRange usedRange = worksheet.RangeUsed();

                        if (usedRange == null)
                            continue;

                        // ------------------------------------------------
                        // Leer encabezados
                        // ------------------------------------------------

                        var headerRow = usedRange.FirstRow();

                        Dictionary<int, string> headers =
                            new Dictionary<int, string>();

                        int elementIdColumn = -1;

                        foreach (IXLCell cell in headerRow.Cells())
                        {
                            string header = cell.Value
                                .ToString()
                                .Trim();

                            if (string.IsNullOrWhiteSpace(header))
                                continue;

                            headers[cell.Address.ColumnNumber] = header;

                            // Aceptamos ElementID y ElementId
                            if (string.Equals(
                                    header,
                                    "ElementID",
                                    StringComparison.OrdinalIgnoreCase) ||
                                string.Equals(
                                    header,
                                    "ElementId",
                                    StringComparison.OrdinalIgnoreCase))
                            {
                                elementIdColumn = cell.Address.ColumnNumber;
                            }
                        }

                        // ------------------------------------------------
                        // Si la hoja no tiene ElementID, la saltamos
                        // ------------------------------------------------

                        if (elementIdColumn == -1)
                        {
                            logErrores.Add(
                                $"Hoja '{worksheet.Name}': no se encontró columna ElementID.");

                            continue;
                        }

                        // =================================================
                        // RECORRER FILAS
                        // =================================================

                        foreach (var row in usedRange.RowsUsed().Skip(1))
                        {
                            try
                            {
                                filasProcesadas++;

                                // -----------------------------------------
                                // Leer ElementID
                                // -----------------------------------------

                                string elementIdText = row
                                    .Cell(elementIdColumn)
                                    .Value
                                    .ToString()
                                    .Trim();

                                if (string.IsNullOrWhiteSpace(elementIdText))
                                    continue;

                                if (!int.TryParse(
                                        elementIdText,
                                        NumberStyles.Integer,
                                        CultureInfo.InvariantCulture,
                                        out int elementIdInteger))
                                {
                                    // Intentar por si Excel lo entrega
                                    // como número decimal
                                    if (!double.TryParse(
                                            elementIdText,
                                            NumberStyles.Any,
                                            CultureInfo.InvariantCulture,
                                            out double tempDouble))
                                    {
                                        errores++;

                                        logErrores.Add(
                                            $"Hoja '{worksheet.Name}', fila {row.RowNumber()}: " +
                                            $"ElementID inválido '{elementIdText}'.");

                                        continue;
                                    }

                                    elementIdInteger =
                                        Convert.ToInt32(tempDouble);
                                }

                                // -----------------------------------------
                                // Obtener elemento Revit
                                // -----------------------------------------

                                ElementId elementId =
                                    new ElementId((long)elementIdInteger);

                                Element element =
                                    doc.GetElement(elementId);

                                if (element == null)
                                {
                                    elementosNoEncontrados++;

                                    logErrores.Add(
                                        $"Hoja '{worksheet.Name}', fila {row.RowNumber()}: " +
                                        $"ElementID {elementIdInteger} no encontrado.");

                                    continue;
                                }

                                elementosEncontrados++;

                                // =================================================
                                // RECORRER COLUMNAS / PARÁMETROS
                                // =================================================

                                foreach (KeyValuePair<int, string> header in headers)
                                {
                                    int columnNumber = header.Key;
                                    string parameterName = header.Value;

                                    // No modificar ElementID
                                    if (columnNumber == elementIdColumn)
                                        continue;

                                    // ---------------------------------------------
                                    // Buscar parámetro por nombre
                                    // ---------------------------------------------

                                    Parameter parameter =
                                        element.LookupParameter(parameterName);

                                    if (parameter == null)
                                    {
                                        parametrosNoEncontrados++;

                                        // NUEVO: guardar detalle
                                        parametrosNoEncontradosLista.Add(
                                            new ParametroNoEncontrado
                                            {
                                                Hoja = worksheet.Name,
                                                Fila = row.RowNumber(),
                                                ElementID = elementIdInteger,
                                                Parametro = parameterName
                                            });

                                        continue;
                                    }

                                    // ---------------------------------------------
                                    // Verificar si se puede escribir
                                    // ---------------------------------------------

                                    if (parameter.IsReadOnly)
                                    {
                                        parametrosSoloLectura++;

                                        logErrores.Add(
                                            $"Elemento {elementIdInteger}: " +
                                            $"parámetro '{parameterName}' es Read Only.");

                                        continue;
                                    }

                                    // ---------------------------------------------
                                    // Valor Excel
                                    // ---------------------------------------------

                                    IXLCell cell =
                                        row.Cell(columnNumber);

                                    if (cell.IsEmpty())
                                        continue;

                                    string value =
                                        cell.Value.ToString().Trim();

                                    if (string.IsNullOrWhiteSpace(value))
                                        continue;

                                    // ---------------------------------------------
                                    // Asignar valor
                                    // ---------------------------------------------

                                    if (SetParameterValue(parameter, value))
                                    {
                                        valoresAsignados++;
                                    }
                                    else
                                    {
                                        errores++;

                                        logErrores.Add(
                                            $"Elemento {elementIdInteger}: " +
                                            $"no se pudo asignar '{value}' " +
                                            $"al parámetro '{parameterName}'.");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                errores++;

                                logErrores.Add(
                                    $"Hoja '{worksheet.Name}', fila {row.RowNumber()}: " +
                                    ex.Message);
                            }
                        }
                    }

                    trans.Commit();
                }
            }

            // ============================================================
            // CREAR REPORTE DE PARAMETROS NO ENCONTRADOS
            // ============================================================

            string rutaReporte =
                CrearReporteParametrosNoEncontrados(
                    parametrosNoEncontradosLista);

            // ============================================================
            // RESUMEN
            // ============================================================

            string resumen =
                "Proceso terminado.\n\n" +

                $"Hojas procesadas: {hojasProcesadas}\n" +
                $"Filas procesadas: {filasProcesadas}\n" +
                $"Elementos encontrados: {elementosEncontrados}\n" +
                $"Valores asignados: {valoresAsignados}\n\n" +

                $"Elementos no encontrados: {elementosNoEncontrados}\n" +
                $"Parámetros no encontrados: {parametrosNoEncontrados}\n" +
                $"Parámetros Read Only: {parametrosSoloLectura}\n" +
                $"Errores: {errores}\n\n" +

                (parametrosNoEncontrados > 0
                    ? "Reporte generado:\n" + rutaReporte
                    : "No hubo parámetros no encontrados.");

            TaskDialog.Show("BimDataSync", resumen);

            return Result.Succeeded;
        }


        // ================================================================
        // CREAR REPORTE DE PARAMETROS NO ENCONTRADOS
        // ================================================================

        private string CrearReporteParametrosNoEncontrados(
            List<ParametroNoEncontrado> lista)
        {
            string carpeta =
                Path.GetDirectoryName(ExcelPath);

            string nombreArchivo =
                Path.GetFileNameWithoutExtension(ExcelPath);

            string rutaReporte =
                Path.Combine(
                    carpeta,
                    nombreArchivo + "_BimDataSync_Log.xlsx");

            // Si no hay parámetros faltantes, no necesitamos generar
            // un archivo de reporte.
            if (lista == null || lista.Count == 0)
                return rutaReporte;

            using (XLWorkbook workbook = new XLWorkbook())
            {
                // ========================================================
                // HOJA 1 - RESUMEN
                // ========================================================

                IXLWorksheet resumen =
                    workbook.Worksheets.Add("Resumen");

                resumen.Cell(1, 1).Value = "Parámetro";
                resumen.Cell(1, 2).Value = "Veces no encontrado";

                var agrupados =
                    lista
                    .GroupBy(x => x.Parametro)
                    .OrderByDescending(x => x.Count());

                int filaResumen = 2;

                foreach (var grupo in agrupados)
                {
                    resumen.Cell(filaResumen, 1).Value =
                        grupo.Key;

                    resumen.Cell(filaResumen, 2).Value =
                        grupo.Count();

                    filaResumen++;
                }

                // Formato encabezados
                var headerResumen =
                    resumen.Range(1, 1, 1, 2);

                headerResumen.Style.Font.Bold = true;
                headerResumen.Style.Fill.BackgroundColor =
                    XLColor.Green;

                headerResumen.Style.Font.FontColor =
                    XLColor.White;

                resumen.Columns().AdjustToContents();


                // ========================================================
                // HOJA 2 - DETALLE
                // ========================================================

                IXLWorksheet detalle =
                    workbook.Worksheets.Add("Detalle");

                detalle.Cell(1, 1).Value = "Hoja";
                detalle.Cell(1, 2).Value = "Fila Excel";
                detalle.Cell(1, 3).Value = "ElementID";
                detalle.Cell(1, 4).Value = "Parámetro";

                int filaDetalle = 2;

                foreach (ParametroNoEncontrado item in lista)
                {
                    detalle.Cell(filaDetalle, 1).Value =
                        item.Hoja;

                    detalle.Cell(filaDetalle, 2).Value =
                        item.Fila;

                    detalle.Cell(filaDetalle, 3).Value =
                        item.ElementID;

                    detalle.Cell(filaDetalle, 4).Value =
                        item.Parametro;

                    filaDetalle++;
                }

                var headerDetalle =
                    detalle.Range(1, 1, 1, 4);

                headerDetalle.Style.Font.Bold = true;
                headerDetalle.Style.Fill.BackgroundColor =
                    XLColor.Green;

                headerDetalle.Style.Font.FontColor =
                    XLColor.White;

                detalle.Columns().AdjustToContents();


                // ========================================================
                // GUARDAR
                // ========================================================

                workbook.SaveAs(rutaReporte);
            }

            return rutaReporte;
        }


        // ================================================================
        // ASIGNAR VALOR AL PARAMETRO
        // ================================================================

        private bool SetParameterValue(
            Parameter parameter,
            string value)
        {
            try
            {
                switch (parameter.StorageType)
                {
                    // ----------------------------------------------------
                    // STRING
                    // ----------------------------------------------------

                    case StorageType.String:

                        parameter.Set(value);
                        return true;


                    // ----------------------------------------------------
                    // INTEGER
                    // ----------------------------------------------------

                    case StorageType.Integer:

                        if (int.TryParse(
                                value,
                                NumberStyles.Integer,
                                CultureInfo.CurrentCulture,
                                out int intValue))
                        {
                            parameter.Set(intValue);
                            return true;
                        }

                        return false;


                    // ----------------------------------------------------
                    // DOUBLE
                    // ----------------------------------------------------

                    case StorageType.Double:

                        double doubleValue;

                        // Primero interpretar usando la configuración
                        // regional de Windows/Excel.
                        if (!double.TryParse(
                                value,
                                NumberStyles.Number,
                                CultureInfo.CurrentCulture,
                                out doubleValue))
                        {
                            // Como respaldo, intentar formato internacional.
                            if (!double.TryParse(
                                    value,
                                    NumberStyles.Number,
                                    CultureInfo.InvariantCulture,
                                    out doubleValue))
                            {
                                return false;
                            }
                        }

                        // Convertir desde la unidad mostrada del parámetro
                        // hacia las unidades internas de Revit.
                        double internalValue =
                            UnitUtils.ConvertToInternalUnits(
                                doubleValue,
                                parameter.GetUnitTypeId());

                        parameter.Set(internalValue);

                        return true;


                    // ----------------------------------------------------
                    // ELEMENT ID
                    // ----------------------------------------------------

                    case StorageType.ElementId:

                        if (long.TryParse(
                                value,
                                NumberStyles.Integer,
                                CultureInfo.CurrentCulture,
                                out long referencedId))
                        {
                            parameter.Set(
                                new ElementId(referencedId));

                            return true;
                        }

                        return false;


                    // ----------------------------------------------------
                    // NONE
                    // ----------------------------------------------------

                    case StorageType.None:
                        return false;


                    default:
                        return false;
                }
            }
            catch
            {
                return false;
            }
        }
    }


    // ====================================================================
    // MODELO PARA REGISTRAR PARAMETROS NO ENCONTRADOS
    // ====================================================================

    public class ParametroNoEncontrado
    {
        public string Hoja { get; set; }
        public int Fila { get; set; }
        public long ElementID { get; set; }
        public string Parametro { get; set; }
    }
}