using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using ClosedXML.Excel;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace ICCU_CrearParametros
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        // ============================================================
        // CONFIGURATION
        // ============================================================

        private const string SharedParameterFileName =
            "ICCU_SharedParameters.txt";


        // ============================================================
        // MAIN COMMAND
        // ============================================================

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            string previousSharedParameterFile = null;

            try
            {
                // ----------------------------------------------------
                // 1. Select Excel file
                // ----------------------------------------------------

                string excelPath = SelectExcelFile();

                if (string.IsNullOrWhiteSpace(excelPath))
                {
                    return Result.Cancelled;
                }


                // ----------------------------------------------------
                // 2. Select worksheet
                // ----------------------------------------------------

                string sheetName = SelectWorksheet(excelPath);

                if (string.IsNullOrWhiteSpace(sheetName))
                {
                    return Result.Cancelled;
                }


                // ----------------------------------------------------
                // 3. Read Excel
                // ----------------------------------------------------

                List<ParameterRow> rows =
                    ReadExcel(excelPath, sheetName);


                if (rows.Count == 0)
                {
                    TaskDialog.Show(
                        "ICCU - Create Parameters",
                        "No valid parameter rows were found.");

                    return Result.Cancelled;
                }


                // ----------------------------------------------------
                // 4. Prepare shared parameter file
                // ----------------------------------------------------

                string sharedParameterPath =
                    Path.Combine(
                        Path.GetDirectoryName(excelPath),
                        SharedParameterFileName);


                previousSharedParameterFile =
                    uiApp.Application.SharedParametersFilename;


                PrepareSharedParameterFile(
                    sharedParameterPath);


                uiApp.Application.SharedParametersFilename =
                    sharedParameterPath;


                DefinitionFile definitionFile =
                    uiApp.Application.OpenSharedParameterFile();


                if (definitionFile == null)
                {
                    TaskDialog.Show(
                        "ICCU - Create Parameters",
                        "Could not open the shared parameter file.");

                    return Result.Failed;
                }


                // ----------------------------------------------------
                // 5. Counters
                // ----------------------------------------------------

                int created = 0;
                int existing = 0;
                int bindingsProcessed = 0;
                int errors = 0;


                List<string> errorMessages =
                    new List<string>();


                // ----------------------------------------------------
                // 6. Process parameters
                // ----------------------------------------------------

                using (Transaction transaction =
                    new Transaction(
                        doc,
                        "ICCU - Create Parameters"))
                {
                    transaction.Start();


                    foreach (ParameterRow row in rows)
                    {
                        try
                        {
                            ProcessParameter(
                                uiApp,
                                doc,
                                definitionFile,
                                row,
                                ref created,
                                ref existing,
                                ref bindingsProcessed);
                        }
                        catch (Exception ex)
                        {
                            errors++;

                            errorMessages.Add(
                                $"Row {row.RowNumber} - {row.ParameterName}\n" +
                                ex.Message);
                        }
                    }


                    transaction.Commit();
                }


                // ----------------------------------------------------
                // 7. Restore previous shared parameter file
                // ----------------------------------------------------

                uiApp.Application.SharedParametersFilename =
                    previousSharedParameterFile;


                // ----------------------------------------------------
                // 8. Result
                // ----------------------------------------------------

                string result =
                    "PARAMETER CREATION COMPLETED\n\n" +
                    $"Rows processed: {rows.Count}\n" +
                    $"New parameters: {created}\n" +
                    $"Existing parameters: {existing}\n" +
                    $"Bindings processed: {bindingsProcessed}\n" +
                    $"Errors: {errors}";


                if (errors > 0)
                {
                    result += "\n\nERRORS:\n";

                    foreach (string error in errorMessages)
                    {
                        result += "\n" + error + "\n";
                    }
                }


                TaskDialog.Show(
                    "ICCU - Create Parameters",
                    result);


                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                // Restore previous shared parameter file
                // in case of an unexpected error.

                if (previousSharedParameterFile != null)
                {
                    try
                    {
                        uiApp.Application.SharedParametersFilename =
                            previousSharedParameterFile;
                    }
                    catch
                    {
                        // Ignore restore error.
                    }
                }


                message = ex.ToString();


                TaskDialog.Show(
                    "ICCU - Create Parameters",
                    ex.Message);


                return Result.Failed;
            }
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
                    "Select parameter matrix";

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
                        "The selected Excel file does not contain worksheets.");
                }


                string sheetList = "";


                for (int i = 0;
                     i < sheetNames.Count;
                     i++)
                {
                    sheetList +=
                        $"{i + 1} - {sheetNames[i]}\r\n";
                }


                string input =
                    Microsoft.VisualBasic.Interaction.InputBox(
                        "Select the worksheet number:\r\n\r\n" +
                        sheetList,
                        "Select worksheet",
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
                        "ICCU - Create Parameters",
                        "Invalid worksheet number.");

                    return null;
                }


                if (selectedNumber < 1 ||
                    selectedNumber > sheetNames.Count)
                {
                    TaskDialog.Show(
                        "ICCU - Create Parameters",
                        "Worksheet number out of range.");

                    return null;
                }


                return sheetNames[selectedNumber - 1];
            }
        }


        // ============================================================
        // READ EXCEL
        // ============================================================

        private List<ParameterRow> ReadExcel(
            string excelPath,
            string sheetName)
        {
            List<ParameterRow> rows =
                new List<ParameterRow>();


            using (
                XLWorkbook workbook =
                    new XLWorkbook(excelPath))
            {
                IXLWorksheet worksheet =
                    workbook.Worksheet(sheetName);


                // ----------------------------------------------------
                // Find headers
                // ----------------------------------------------------

                IXLRow headerRow =
                    worksheet.Row(1);


                Dictionary<string, int> columns =
                    new Dictionary<string, int>(
                        StringComparer.OrdinalIgnoreCase);


                foreach (IXLCell cell in
                    headerRow.CellsUsed())
                {
                    string header =
                        cell.GetString().Trim();


                    if (!string.IsNullOrWhiteSpace(header))
                    {
                        columns[header] =
                            cell.Address.ColumnNumber;
                    }
                }


                // ----------------------------------------------------
                // Required headers
                // ----------------------------------------------------

                string[] requiredHeaders =
                {
                    "Grupo",
                    "Categorias",
                    "Parametros",
                    "Tipo",
                    "Binding"
                };


                foreach (string header in
                    requiredHeaders)
                {
                    if (!columns.ContainsKey(header))
                    {
                        throw new Exception(
                            $"Required column not found: {header}");
                    }
                }


                // ----------------------------------------------------
                // Read rows
                // ----------------------------------------------------

                IXLRow lastUsedRow =
                    worksheet.LastRowUsed();


                if (lastUsedRow == null)
                {
                    return rows;
                }


                int lastRow =
                    lastUsedRow.RowNumber();


                for (int rowNumber = 2;
                     rowNumber <= lastRow;
                     rowNumber++)
                {
                    string group =
                        worksheet.Cell(
                            rowNumber,
                            columns["Grupo"])
                        .GetString()
                        .Trim();


                    string categories =
                        worksheet.Cell(
                            rowNumber,
                            columns["Categorias"])
                        .GetString()
                        .Trim();


                    string parameterName =
                        worksheet.Cell(
                            rowNumber,
                            columns["Parametros"])
                        .GetString()
                        .Trim();


                    string type =
                        worksheet.Cell(
                            rowNumber,
                            columns["Tipo"])
                        .GetString()
                        .Trim();


                    string binding =
                        worksheet.Cell(
                            rowNumber,
                            columns["Binding"])
                        .GetString()
                        .Trim();


                    // Ignore completely empty rows

                    if (string.IsNullOrWhiteSpace(
                        parameterName))
                    {
                        continue;
                    }


                    rows.Add(
                        new ParameterRow
                        {
                            RowNumber = rowNumber,
                            Group = group,
                            Categories = categories,
                            ParameterName = parameterName,
                            Type = type,
                            Binding = binding
                        });
                }
            }


            return rows;
        }


        // ============================================================
        // PREPARE SHARED PARAMETER FILE
        // ============================================================

        private void PrepareSharedParameterFile(
            string path)
        {
            if (File.Exists(path))
            {
                return;
            }


            string directory =
                Path.GetDirectoryName(path);


            if (!Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }


            File.WriteAllText(
                path,
                "# This is a Revit shared parameter file.\r\n" +
                "# Do not edit manually.\r\n" +
                "*META\tVERSION\tMINVERSION\r\n" +
                "META\t2\t1\r\n" +
                "*GROUP\tID\tNAME\r\n");
        }


        // ============================================================
        // PROCESS PARAMETER
        // ============================================================

        private void ProcessParameter(
            UIApplication uiApp,
            Document doc,
            DefinitionFile definitionFile,
            ParameterRow row,
            ref int created,
            ref int existing,
            ref int bindingsProcessed)
        {
            // --------------------------------------------------------
            // 1. Resolve data type
            // --------------------------------------------------------

            ForgeTypeId dataType =
                GetDataType(row.Type);


            // --------------------------------------------------------
            // 2. Resolve parameter group
            // --------------------------------------------------------

            ForgeTypeId groupTypeId =
                GetGroupTypeId(row.Group);


            // --------------------------------------------------------
            // 3. Get or create definition
            // --------------------------------------------------------

            DefinitionGroup definitionGroup =
                GetOrCreateDefinitionGroup(
                    definitionFile,
                    "ICCU");


            ExternalDefinition definition =
                GetExistingDefinition(
                    definitionGroup,
                    row.ParameterName);


            if (definition == null)
            {
                ExternalDefinitionCreationOptions options =
                    new ExternalDefinitionCreationOptions(
                        row.ParameterName,
                        dataType);


                definition =
                    definitionGroup.Definitions.Create(options)
                    as ExternalDefinition;


                created++;
            }
            else
            {
                existing++;
            }


            // --------------------------------------------------------
            // 4. Resolve categories
            // --------------------------------------------------------

            CategorySet categorySet =
                uiApp.Application.Create.NewCategorySet();


            string[] categoryNames =
                row.Categories
                    .Split(',')
                    .Select(x => x.Trim())
                    .Where(x =>
                        !string.IsNullOrWhiteSpace(x))
                    .ToArray();


            foreach (string categoryName
                in categoryNames)
            {
                Category category =
                    FindCategory(
                        doc,
                        categoryName);


                if (category == null)
                {
                    throw new Exception(
                        $"Category not found: {categoryName}");
                }


                if (!category.AllowsBoundParameters)
                {
                    throw new Exception(
                        $"Category does not allow bound parameters: {categoryName}");
                }


                categorySet.Insert(category);
            }


            // --------------------------------------------------------
            // 5. Create binding
            // --------------------------------------------------------

            Binding binding;


            if (row.Binding.Equals(
                    "Instance",
                    StringComparison.OrdinalIgnoreCase))
            {
                binding =
                    uiApp.Application.Create
                        .NewInstanceBinding(
                            categorySet);
            }
            else if (row.Binding.Equals(
                        "Type",
                        StringComparison.OrdinalIgnoreCase))
            {
                binding =
                    uiApp.Application.Create
                        .NewTypeBinding(
                            categorySet);
            }
            else
            {
                throw new Exception(
                    $"Invalid Binding value: {row.Binding}. " +
                    "Use Instance or Type.");
            }


            // --------------------------------------------------------
            // 6. Insert / update binding
            // --------------------------------------------------------

            BindingMap bindingMap =
                doc.ParameterBindings;


            bool inserted =
                bindingMap.Insert(
                    definition,
                    binding,
                    groupTypeId);


            if (inserted)
            {
                bindingsProcessed++;
            }
            else
            {
                bool reinserted =
                    bindingMap.ReInsert(
                        definition,
                        binding,
                        groupTypeId);


                if (reinserted)
                {
                    bindingsProcessed++;
                }
                else
                {
                    throw new Exception(
                        "Could not insert or update the parameter binding.");
                }
            }
        }


        // ============================================================
        // DEFINITION GROUP
        // ============================================================

        private DefinitionGroup GetOrCreateDefinitionGroup(
            DefinitionFile definitionFile,
            string groupName)
        {
            DefinitionGroup group =
                definitionFile.Groups
                    .get_Item(groupName);


            if (group != null)
            {
                return group;
            }


            return definitionFile.Groups.Create(
                groupName);
        }


        // ============================================================
        // EXISTING DEFINITION
        // ============================================================

        private ExternalDefinition GetExistingDefinition(
            DefinitionGroup group,
            string parameterName)
        {
            Definition definition =
                group.Definitions
                    .get_Item(parameterName);


            return definition as ExternalDefinition;
        }


        // ============================================================
        // CATEGORY RESOLVER
        // ============================================================

        private Category FindCategory(
            Document doc,
            string categoryName)
        {
            foreach (Category category
                in doc.Settings.Categories)
            {
                if (category.Name.Equals(
                        categoryName,
                        StringComparison.OrdinalIgnoreCase))
                {
                    return category;
                }
            }


            return null;
        }


        // ============================================================
        // DATA TYPE RESOLVER
        // ============================================================

        private ForgeTypeId GetDataType(
            string type)
        {
            string cleanType =
                new string(
                    type
                        .Where(c =>
                            (c >= 'A' && c <= 'Z') ||
                            (c >= 'a' && c <= 'z'))
                        .ToArray())
                    .ToLowerInvariant();


            if (cleanType == "text")
            {
                return SpecTypeId.String.Text;
            }


            if (cleanType == "integer")
            {
                return SpecTypeId.Int.Integer;
            }


            if (cleanType == "number")
            {
                return SpecTypeId.Number;
            }


            if (cleanType == "yesno")
            {
                return SpecTypeId.Boolean.YesNo;
            }


            if (cleanType.Contains("length"))
            {
                return SpecTypeId.Length;
            }


            if (cleanType == "area")
            {
                return SpecTypeId.Area;
            }


            if (cleanType == "volume")
            {
                return SpecTypeId.Volume;
            }


            if (cleanType == "angle")
            {
                return SpecTypeId.Angle;
            }


            if (cleanType == "mass")
            {
                return SpecTypeId.Mass;
            }


            if (cleanType == "time")
            {
                return SpecTypeId.Time;
            }


            throw new Exception(
                $"Unsupported parameter type: [{type}]");
        }


        // ============================================================
        // GROUP RESOLVER
        // ============================================================

        private ForgeTypeId GetGroupTypeId(
            string group)
        {
            switch (group.Trim())
            {
                case "Construction":
                    return GroupTypeId.Construction;

                case "Constraints":
                    return GroupTypeId.Constraints;

                case "Dimensions":
                    return GroupTypeId.Constraints;

                case "Graphics":
                    return GroupTypeId.Graphics;

                case "Identity Data":
                    return GroupTypeId.IdentityData;

                case "Materials and Finishes":
                    return GroupTypeId.Materials;

                case "Phasing":
                    return GroupTypeId.Phasing;

                case "Structural":
                    return GroupTypeId.Structural;

                case "Text":
                    return GroupTypeId.Text;

                case "Visibility":
                    return GroupTypeId.Visibility;

                default:
                    throw new Exception(
                        $"Unsupported parameter group: {group}");
            }
        }
    }


    // ================================================================
    // EXCEL ROW MODEL
    // ================================================================

    public class ParameterRow
    {
        public int RowNumber { get; set; }

        public string Group { get; set; }

        public string Categories { get; set; }

        public string ParameterName { get; set; }

        public string Type { get; set; }

        public string Binding { get; set; }
    }
}