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
        // CONFIGURACION
        // ============================================================
        // La matriz Excel y la hoja se seleccionan al ejecutar el comando.


        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // --------------------------------------------------------
            // 1. Seleccionar matriz Excel
            // --------------------------------------------------------

            string excelPath = SelectExcelFile();

            if (string.IsNullOrWhiteSpace(excelPath))
                return Result.Cancelled;

            // --------------------------------------------------------
            // 2. Seleccionar hoja
            // --------------------------------------------------------

            string sheetName = SelectWorksheet(excelPath);

            if (string.IsNullOrWhiteSpace(sheetName))
                return Result.Cancelled;

            int hojasProcesadas = 0;
            int filasProcesadas = 0;
            int elementosEncontrados = 0;
            int valoresAsignados = 0;
            int parametrosNoEncontrados = 0;
            int elementosNoEncontrados = 0;
            int parametrosSoloLectura = 0;
            int errores = 0;

            List<string> logErrores = new List<string>();

            List<ParametroNoEncontrado> parametrosNoEncontradosLista =
                new List<ParametroNoEncontrado>();

            // --------------------------------------------------------
            // Abrir Excel
            // --------------------------------------------------------

            using (XLWorkbook workbook = new XLWorkbook(excelPath))
            {
                using (Transaction trans = new Transaction(
                    doc,
                    "BimDataSync - Asignar datos"))
                {
                    trans.Start();

                    // =================================================
                    // PROCESAR LA HOJA SELECCIONADA
                    // =================================================

                    IXLWorksheet worksheet =
                        workbook.Worksheet(sheetName);

                    hojasProcesadas++;

                    IXLRange usedRange = worksheet.RangeUsed();

                    if (usedRange == null)
                    {
                        trans.Commit();

                        TaskDialog.Show(
                            "BimDataSync",
                            "La hoja seleccionada no contiene datos.");

                        return Result.Succeeded;
                    }

                    // ------------------------------------------------
                    // Leer encabezados
                    // ------------------------------------------------

                    var headerRow = usedRange.FirstRow();

                    Dictionary<int, string> headers =
                        new Dictionary<int, string>();

                    int elementIdColumn = -1;

                    foreach (IXLCell cell in headerRow.Cells())
                    {
                        string header = cell.Value.ToString().Trim();

                        if (string.IsNullOrWhiteSpace(header))
                            continue;

                        headers[cell.Address.ColumnNumber] = header;

                        if (string.Equals(
                                header,
                                "ElementID",
                                StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(
                                header,
                                "ElementId",
                                StringComparison.OrdinalIgnoreCase))
                        {
                            elementIdColumn =
                                cell.Address.ColumnNumber;
                        }
                    }

                    if (elementIdColumn == -1)
                    {
                        logErrores.Add(
                            $"Hoja '{worksheet.Name}': no se encontró columna ElementID.");
                    }
                    else
                    {
                        // =================================================
                        // RECORRER FILAS
                        // =================================================

                        foreach (var row in usedRange.RowsUsed().Skip(1))
                        {
                            try
                            {
                                filasProcesadas++;

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

                                    if (columnNumber == elementIdColumn)
                                        continue;

                                    Parameter parameter =
                                        element.LookupParameter(parameterName);

                                    if (parameter == null)
                                    {
                                        parametrosNoEncontrados++;

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

                                    if (parameter.IsReadOnly)
                                    {
                                        parametrosSoloLectura++;

                                        logErrores.Add(
                                            $"Elemento {elementIdInteger}: " +
                                            $"parámetro '{parameterName}' es Read Only.");

                                        continue;
                                    }

                                    IXLCell cell =
                                        row.Cell(columnNumber);

                                    if (cell.IsEmpty())
                                        continue;

                                    string value =
                                        cell.Value.ToString().Trim();

                                    if (string.IsNullOrWhiteSpace(value))
                                        continue;

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
                    excelPath,
                    parametrosNoEncontradosLista);

            // ============================================================
            // RESUMEN
            // ============================================================

            string resumen =
                "Proceso terminado.\n\n" +

                $"Matriz: {Path.GetFileName(excelPath)}\n" +
                $"Hoja: {sheetName}\n\n" +

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


        // ============================================================
        // SELECT EXCEL FILE
        // ============================================================

        private string SelectExcelFile()
        {
            using (
                System.Windows.Forms.OpenFileDialog dialog =
                    new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title =
                    "Seleccionar matriz Excel";

                dialog.Filter =
                    "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";

                dialog.Multiselect = false;

                if (dialog.ShowDialog() ==
                    System.Windows.Forms.DialogResult.OK)
                {
                    return dialog.FileName;
                }
            }

            return null;
        }


        // ============================================================
        // SELECT WORKSHEET
        // ============================================================

        private string SelectWorksheet(
            string excelPath)
        {
            using (
                XLWorkbook workbook =
                    new XLWorkbook(excelPath))
            {
                List<string> sheetNames =
                    workbook.Worksheets
                        .Select(ws => ws.Name)
                        .ToList();

                if (sheetNames.Count == 0)
                {
                    throw new Exception(
                        "El archivo Excel seleccionado no contiene hojas.");
                }

                string sheetList = "";

                for (
                    int i = 0;
                    i < sheetNames.Count;
                    i++)
                {
                    sheetList +=
                        $"{i + 1} - {sheetNames[i]}\r\n";
                }

                string input =
                    Microsoft.VisualBasic.Interaction.InputBox(
                        "Seleccione el número de la hoja:\r\n\r\n" +
                        sheetList,
                        "Seleccionar hoja",
                        "1");

                if (string.IsNullOrWhiteSpace(input))
                {
                    return null;
                }

                int selectedNumber;

                if (!int.TryParse(
                        input,
                        out selectedNumber))
                {
                    TaskDialog.Show(
                        "BimDataSync",
                        "Número de hoja inválido.");

                    return null;
                }

                if (
                    selectedNumber < 1 ||
                    selectedNumber > sheetNames.Count)
                {
                    TaskDialog.Show(
                        "BimDataSync",
                        "El número de hoja está fuera de rango.");

                    return null;
                }

                return sheetNames[selectedNumber - 1];
            }
        }


        // ================================================================
        // CREAR REPORTE DE PARAMETROS NO ENCONTRADOS
        // ================================================================

        private string CrearReporteParametrosNoEncontrados(
            string excelPath,
            List<ParametroNoEncontrado> lista)
        {
            string carpeta =
                Path.GetDirectoryName(excelPath);

            string nombreArchivo =
                Path.GetFileNameWithoutExtension(excelPath);

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
