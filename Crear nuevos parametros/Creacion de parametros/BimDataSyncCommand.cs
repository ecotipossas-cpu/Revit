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
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            // ============================================================
            // 1. SELECCIONAR MATRIZ EXCEL
            // ============================================================

            string excelPath = SelectExcelFile();

            if (string.IsNullOrWhiteSpace(excelPath))
                return Result.Cancelled;

            // ============================================================
            // 2. SELECCIONAR HOJA
            // ============================================================

            string sheetName = SelectWorksheet(excelPath);

            if (string.IsNullOrWhiteSpace(sheetName))
                return Result.Cancelled;

            int hojasProcesadas = 0;
            int filasProcesadas = 0;
            int elementosEncontrados = 0;
            int valoresProcesados = 0;
            int valoresAsignados = 0;
            int elementosNoEncontrados = 0;
            int parametrosNoEncontrados = 0;
            int parametrosSoloLectura = 0;
            int errores = 0;

            List<string> logErrores = new List<string>();

            // Lista de parámetros realmente no encontrados.
            // SOLO se registra cuando la celda Excel tiene un valor.
            List<ParametroNoEncontrado> parametrosNoEncontradosLista =
                new List<ParametroNoEncontrado>();

            // Registro completo de cada intento de sincronización.
            List<RegistroSincronizacion> registros =
                new List<RegistroSincronizacion>();

            try
            {
                using (XLWorkbook workbook = new XLWorkbook(excelPath))
                {
                    IXLWorksheet worksheet = workbook.Worksheet(sheetName);
                    hojasProcesadas = 1;

                    IXLRange usedRange = worksheet.RangeUsed();

                    if (usedRange == null)
                    {
                        TaskDialog.Show(
                            "BimDataSync",
                            "La hoja seleccionada no contiene datos.");

                        return Result.Failed;
                    }

                    // ========================================================
                    // LEER ENCABEZADOS
                    // ========================================================

                    IXLRangeRow headerRow = usedRange.FirstRow();

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

                    if (elementIdColumn == -1)
                    {
                        TaskDialog.Show(
                            "BimDataSync",
                            "La hoja seleccionada no contiene una columna ElementID.");

                        return Result.Failed;
                    }

                    // ========================================================
                    // TRANSACTION
                    // ========================================================

                    using (Transaction trans = new Transaction(
                        doc,
                        "BimDataSync - Asignar datos"))
                    {
                        trans.Start();

                        // ====================================================
                        // RECORRER FILAS
                        // ====================================================

                        foreach (IXLRangeRow row in usedRange.RowsUsed().Skip(1))
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

                                long elementIdInteger;

                                if (!long.TryParse(
                                        elementIdText,
                                        NumberStyles.Integer,
                                        CultureInfo.InvariantCulture,
                                        out elementIdInteger))
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

                                    elementIdInteger = Convert.ToInt64(tempDouble);
                                }

                                ElementId elementId =
                                    new ElementId(elementIdInteger);

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

                                // ====================================================
                                // RECORRER COLUMNAS / PARAMETROS
                                // IMPORTANTE:
                                // Primero se lee el valor Excel.
                                // Si está vacío, NO se busca ni se reporta el parámetro.
                                // ====================================================

                                foreach (KeyValuePair<int, string> header in headers)
                                {
                                    int columnNumber = header.Key;
                                    string parameterName = header.Value;

                                    if (columnNumber == elementIdColumn)
                                        continue;

                                    IXLCell cell = row.Cell(columnNumber);

                                    if (cell.IsEmpty())
                                        continue;

                                    string value = cell.Value
                                        .ToString()
                                        .Trim();

                                    if (string.IsNullOrWhiteSpace(value))
                                        continue;

                                    valoresProcesados++;

                                    // ---------------------------------------------
                                    // Buscar parámetro
                                    // ---------------------------------------------

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
                                                Parametro = parameterName,
                                                ValorExcel = value
                                            });

                                        registros.Add(
                                            new RegistroSincronizacion
                                            {
                                                Hoja = worksheet.Name,
                                                Fila = row.RowNumber(),
                                                ElementID = elementIdInteger,
                                                Parametro = parameterName,
                                                ValorExcel = value,
                                                Estado = "NO ENCONTRADO",
                                                Detalle = "El parámetro no existe en el elemento."
                                            });

                                        continue;
                                    }

                                    // ---------------------------------------------
                                    // Read Only
                                    // ---------------------------------------------

                                    if (parameter.IsReadOnly)
                                    {
                                        parametrosSoloLectura++;

                                        registros.Add(
                                            new RegistroSincronizacion
                                            {
                                                Hoja = worksheet.Name,
                                                Fila = row.RowNumber(),
                                                ElementID = elementIdInteger,
                                                Parametro = parameterName,
                                                ValorExcel = value,
                                                Estado = "READ ONLY",
                                                Detalle = "El parámetro existe pero es de solo lectura."
                                            });

                                        continue;
                                    }

                                    // ---------------------------------------------
                                    // Asignar valor
                                    // ---------------------------------------------

                                    string detalleAsignacion;

                                    if (SetParameterValue(
                                            parameter,
                                            value,
                                            out detalleAsignacion))
                                    {
                                        valoresAsignados++;

                                        registros.Add(
                                            new RegistroSincronizacion
                                            {
                                                Hoja = worksheet.Name,
                                                Fila = row.RowNumber(),
                                                ElementID = elementIdInteger,
                                                Parametro = parameterName,
                                                ValorExcel = value,
                                                Estado = "ASIGNADO",
                                                Detalle = detalleAsignacion
                                            });
                                    }
                                    else
                                    {
                                        errores++;

                                        registros.Add(
                                            new RegistroSincronizacion
                                            {
                                                Hoja = worksheet.Name,
                                                Fila = row.RowNumber(),
                                                ElementID = elementIdInteger,
                                                Parametro = parameterName,
                                                ValorExcel = value,
                                                Estado = "ERROR ASIGNACION",
                                                Detalle = detalleAsignacion
                                            });

                                        logErrores.Add(
                                            $"Hoja '{worksheet.Name}', fila {row.RowNumber()}, " +
                                            $"ElementID {elementIdInteger}: " +
                                            $"no se pudo asignar '{value}' al parámetro " +
                                            $"'{parameterName}'. {detalleAsignacion}");
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

                        trans.Commit();
                    }
                }
            }
            catch (Exception ex)
            {
                TaskDialog.Show(
                    "BimDataSync",
                    "Se produjo un error durante el proceso:\n\n" +
                    ex.Message);

                return Result.Failed;
            }

            // ============================================================
            // REPORTE
            // ============================================================

            string rutaReporte =
                CrearReporte(
                    excelPath,
                    sheetName,
                    parametrosNoEncontradosLista,
                    registros,
                    logErrores);

            // ============================================================
            // RESUMEN
            // ============================================================

            string resumen =
                "Proceso terminado.\n\n" +
                $"Matriz: {Path.GetFileName(excelPath)}\n" +
                $"Hoja: {sheetName}\n\n" +
                $"Filas procesadas: {filasProcesadas}\n" +
                $"Elementos encontrados: {elementosEncontrados}\n" +
                $"Valores con dato en Excel: {valoresProcesados}\n" +
                $"Valores asignados: {valoresAsignados}\n\n" +
                $"Elementos no encontrados: {elementosNoEncontrados}\n" +
                $"Parámetros no encontrados: {parametrosNoEncontrados}\n" +
                $"Parámetros Read Only: {parametrosSoloLectura}\n" +
                $"Errores de asignación: {errores}\n\n" +
                "Reporte generado:\n" + rutaReporte;

            TaskDialog.Show("BimDataSync", resumen);

            return Result.Succeeded;
        }

        // ============================================================
        // SELECCIONAR EXCEL
        // ============================================================

        private string SelectExcelFile()
        {
            using (System.Windows.Forms.OpenFileDialog dialog =
                new System.Windows.Forms.OpenFileDialog())
            {
                dialog.Title = "Seleccionar matriz Excel";
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
        // SELECCIONAR HOJA
        // ============================================================

        private string SelectWorksheet(string excelPath)
        {
            using (XLWorkbook workbook = new XLWorkbook(excelPath))
            {
                List<string> sheetNames = workbook.Worksheets
                    .Select(ws => ws.Name)
                    .ToList();

                if (sheetNames.Count == 0)
                {
                    TaskDialog.Show(
                        "BimDataSync",
                        "El archivo Excel seleccionado no contiene hojas.");

                    return null;
                }

                string sheetList = "";

                for (int i = 0; i < sheetNames.Count; i++)
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
                    return null;

                if (!int.TryParse(input, out int selectedNumber))
                {
                    TaskDialog.Show(
                        "BimDataSync",
                        "Número de hoja inválido.");

                    return null;
                }

                if (selectedNumber < 1 ||
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

        // ============================================================
        // ASIGNAR VALOR AL PARAMETRO
        // ============================================================

        private bool SetParameterValue(
            Parameter parameter,
            string value,
            out string detalle)
        {
            detalle = "";

            try
            {
                switch (parameter.StorageType)
                {
                    // ----------------------------------------------------
                    // STRING
                    // ----------------------------------------------------

                    case StorageType.String:

                        parameter.Set(value);
                        detalle = "Texto asignado correctamente.";
                        return true;

                    // ----------------------------------------------------
                    // INTEGER
                    // ----------------------------------------------------

                    case StorageType.Integer:

                        if (int.TryParse(
                                value,
                                NumberStyles.Integer,
                                CultureInfo.CurrentCulture,
                                out int intValue) ||
                            int.TryParse(
                                value,
                                NumberStyles.Integer,
                                CultureInfo.InvariantCulture,
                                out intValue))
                        {
                            parameter.Set(intValue);
                            detalle = "Entero asignado correctamente.";
                            return true;
                        }

                        detalle =
                            "El valor no pudo convertirse a Integer.";
                        return false;

                    // ----------------------------------------------------
                    // DOUBLE
                    // ----------------------------------------------------

                    case StorageType.Double:

                        if (!TryParseDouble(value, out double doubleValue))
                        {
                            detalle =
                                "El valor no pudo convertirse a número.";
                            return false;
                        }

                        // ------------------------------------------------
                        // IMPORTANTE:
                        // Los parámetros BIM_Quantities de la matriz fueron
                        // creados como tipo Number. En ese caso NO debemos
                        // convertir el valor a pies internos de Revit.
                        // 92.33 debe quedar como 92.33.
                        // ------------------------------------------------

                        // En Revit 2024, "Number" es un SPEC, no un UnitTypeId.
                        // Por eso se debe comprobar mediante Definition.GetDataType().
                        ForgeTypeId dataType =
                            parameter.Definition.GetDataType();

                        if (dataType == SpecTypeId.Number)
                        {
                            parameter.Set(doubleValue);
                            detalle =
                                "Number: valor asignado directamente sin conversión de unidades.";
                            return true;
                        }

                        // Para parámetros que realmente tienen una unidad
                        // Revit (Length, Area, Volume, etc.), sí convertir.
                        ForgeTypeId unitTypeId =
                            parameter.GetUnitTypeId();

                        double internalValue =
                            UnitUtils.ConvertToInternalUnits(
                                doubleValue,
                                unitTypeId);

                        parameter.Set(internalValue);

                        detalle =
                            "Valor numérico convertido a unidades internas de Revit.";

                        return true;

                    // ----------------------------------------------------
                    // ELEMENT ID
                    // ----------------------------------------------------

                    case StorageType.ElementId:

                        if (long.TryParse(
                                value,
                                NumberStyles.Integer,
                                CultureInfo.CurrentCulture,
                                out long referencedId) ||
                            long.TryParse(
                                value,
                                NumberStyles.Integer,
                                CultureInfo.InvariantCulture,
                                out referencedId))
                        {
                            parameter.Set(
                                new ElementId(referencedId));

                            detalle =
                                "ElementId asignado correctamente.";

                            return true;
                        }

                        detalle =
                            "El valor no pudo convertirse a ElementId.";
                        return false;

                    case StorageType.None:
                        detalle = "StorageType.None no es asignable.";
                        return false;

                    default:
                        detalle =
                            "StorageType no soportado por BimDataSync.";
                        return false;
                }
            }
            catch (Exception ex)
            {
                detalle = ex.Message;
                return false;
            }
        }

        // ============================================================
        // CONVERTIR TEXTO A DOUBLE
        // ============================================================

        private bool TryParseDouble(
            string value,
            out double result)
        {
            // Primero configuración regional de Windows/Excel.
            if (double.TryParse(
                    value,
                    NumberStyles.Number,
                    CultureInfo.CurrentCulture,
                    out result))
            {
                return true;
            }

            // Respaldo internacional.
            if (double.TryParse(
                    value,
                    NumberStyles.Number,
                    CultureInfo.InvariantCulture,
                    out result))
            {
                return true;
            }

            result = 0;
            return false;
        }

        // ============================================================
        // CREAR REPORTE
        // ============================================================

        private string CrearReporte(
            string excelPath,
            string sheetName,
            List<ParametroNoEncontrado> parametrosNoEncontradosLista,
            List<RegistroSincronizacion> registros,
            List<string> logErrores)
        {
            string carpeta =
                Path.GetDirectoryName(excelPath);

            string nombreArchivo =
                Path.GetFileNameWithoutExtension(excelPath);

            string rutaReporte =
                Path.Combine(
                    carpeta,
                    nombreArchivo + "_BimDataSync_Log.xlsx");

            using (XLWorkbook workbook = new XLWorkbook())
            {
                // ========================================================
                // HOJA 1 - RESUMEN DE PARAMETROS NO ENCONTRADOS
                // ========================================================

                IXLWorksheet resumen =
                    workbook.Worksheets.Add("Resumen");

                resumen.Cell(1, 1).Value = "Parámetro";
                resumen.Cell(1, 2).Value = "Veces no encontrado";

                var agrupados =
                    parametrosNoEncontradosLista
                    .GroupBy(x => x.Parametro)
                    .OrderByDescending(x => x.Count());

                int filaResumen = 2;

                foreach (var grupo in agrupados)
                {
                    resumen.Cell(filaResumen, 1).Value = grupo.Key;
                    resumen.Cell(filaResumen, 2).Value = grupo.Count();
                    filaResumen++;
                }

                FormatearEncabezado(
                    resumen.Range(1, 1, 1, 2));

                resumen.Columns().AdjustToContents();

                // ========================================================
                // HOJA 2 - DETALLE DE PARAMETROS NO ENCONTRADOS
                // ========================================================

                IXLWorksheet detalle =
                    workbook.Worksheets.Add("Detalle");

                detalle.Cell(1, 1).Value = "Hoja";
                detalle.Cell(1, 2).Value = "Fila Excel";
                detalle.Cell(1, 3).Value = "ElementID";
                detalle.Cell(1, 4).Value = "Parámetro";
                detalle.Cell(1, 5).Value = "Valor Excel";

                int filaDetalle = 2;

                foreach (ParametroNoEncontrado item in
                         parametrosNoEncontradosLista)
                {
                    detalle.Cell(filaDetalle, 1).Value = item.Hoja;
                    detalle.Cell(filaDetalle, 2).Value = item.Fila;
                    detalle.Cell(filaDetalle, 3).Value = item.ElementID;
                    detalle.Cell(filaDetalle, 4).Value = item.Parametro;
                    detalle.Cell(filaDetalle, 5).Value = item.ValorExcel;
                    filaDetalle++;
                }

                FormatearEncabezado(
                    detalle.Range(1, 1, 1, 5));

                detalle.Columns().AdjustToContents();

                // ========================================================
                // HOJA 3 - SINCRONIZACION
                // ========================================================

                IXLWorksheet sincronizacion =
                    workbook.Worksheets.Add("Sincronizacion");

                sincronizacion.Cell(1, 1).Value = "Hoja";
                sincronizacion.Cell(1, 2).Value = "Fila Excel";
                sincronizacion.Cell(1, 3).Value = "ElementID";
                sincronizacion.Cell(1, 4).Value = "Parámetro";
                sincronizacion.Cell(1, 5).Value = "Valor Excel";
                sincronizacion.Cell(1, 6).Value = "Estado";
                sincronizacion.Cell(1, 7).Value = "Detalle";

                int filaSync = 2;

                foreach (RegistroSincronizacion item in registros)
                {
                    sincronizacion.Cell(filaSync, 1).Value = item.Hoja;
                    sincronizacion.Cell(filaSync, 2).Value = item.Fila;
                    sincronizacion.Cell(filaSync, 3).Value = item.ElementID;
                    sincronizacion.Cell(filaSync, 4).Value = item.Parametro;
                    sincronizacion.Cell(filaSync, 5).Value = item.ValorExcel;
                    sincronizacion.Cell(filaSync, 6).Value = item.Estado;
                    sincronizacion.Cell(filaSync, 7).Value = item.Detalle;
                    filaSync++;
                }

                FormatearEncabezado(
                    sincronizacion.Range(1, 1, 1, 7));

                sincronizacion.Columns().AdjustToContents();

                // ========================================================
                // HOJA 4 - ERRORES GENERALES
                // ========================================================

                IXLWorksheet errores =
                    workbook.Worksheets.Add("Errores");

                errores.Cell(1, 1).Value = "Detalle";

                int filaError = 2;

                foreach (string error in logErrores)
                {
                    errores.Cell(filaError, 1).Value = error;
                    filaError++;
                }

                FormatearEncabezado(
                    errores.Range(1, 1, 1, 1));

                errores.Columns().AdjustToContents();

                // ========================================================
                // GUARDAR
                // ========================================================

                workbook.SaveAs(rutaReporte);
            }

            return rutaReporte;
        }

        // ============================================================
        // FORMATO DE ENCABEZADOS
        // ============================================================

        private void FormatearEncabezado(IXLRange range)
        {
            range.Style.Font.Bold = true;
            range.Style.Fill.BackgroundColor = XLColor.Green;
            range.Style.Font.FontColor = XLColor.White;
        }
    }

    // ====================================================================
    // MODELO - PARAMETRO NO ENCONTRADO
    // ====================================================================

    public class ParametroNoEncontrado
    {
        public string Hoja { get; set; }
        public int Fila { get; set; }
        public long ElementID { get; set; }
        public string Parametro { get; set; }
        public string ValorExcel { get; set; }
    }

    // ====================================================================
    // MODELO - REGISTRO DE SINCRONIZACION
    // ====================================================================

    public class RegistroSincronizacion
    {
        public string Hoja { get; set; }
        public int Fila { get; set; }
        public long ElementID { get; set; }
        public string Parametro { get; set; }
        public string ValorExcel { get; set; }
        public string Estado { get; set; }
        public string Detalle { get; set; }
    }
}
