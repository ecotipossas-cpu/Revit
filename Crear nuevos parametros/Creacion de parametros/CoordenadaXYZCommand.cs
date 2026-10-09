using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;

namespace ICCU_Coordenadas
{
    [Transaction(TransactionMode.Manual)]
    public class CoordenadaXYZCommand : IExternalCommand
    {
        private const string PARAMETRO_NOMBRE = "Coordenada (XYZ)";
        private const string GRUPO_PARAMETRO = "ICCU";

        // ================================================================
        // EJECUCION PRINCIPAL
        // ================================================================

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIApplication uiApp = commandData.Application;
            UIDocument uiDoc = uiApp.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // ========================================================
                // 1. SELECCION MULTIPLE
                // ========================================================

                IList<Reference> referencias =
                    uiDoc.Selection.PickObjects(
                        ObjectType.Element,
                        "Seleccione los elementos. Presione ENTER para terminar.");

                if (referencias == null || referencias.Count == 0)
                {
                    return Result.Cancelled;
                }

                // ========================================================
                // 2. OBTENER ELEMENTOS
                // ========================================================

                List<Element> elementosSeleccionados =
                    new List<Element>();

                foreach (Reference referencia in referencias)
                {
                    Element elemento =
                        doc.GetElement(referencia);

                    if (elemento != null)
                    {
                        elementosSeleccionados.Add(elemento);
                    }
                }

                if (elementosSeleccionados.Count == 0)
                {
                    TaskDialog.Show(
                        "Coordenada XYZ",
                        "No se encontraron elementos válidos.");

                    return Result.Cancelled;
                }

                // ========================================================
                // 3. TRANSACCION
                // ========================================================

                int procesados = 0;
                int sinUbicacion = 0;
                int sinParametro = 0;

                using (Transaction trans = new Transaction(
                    doc,
                    "Actualizar Coordenadas XYZ"))
                {
                    trans.Start();

                    // ----------------------------------------------------
                    // PRIMERO: ASEGURAR QUE EL PARAMETRO EXISTE PARA
                    // TODAS LAS CATEGORIAS NECESARIAS
                    // ----------------------------------------------------

                    foreach (Element elemento in elementosSeleccionados)
                    {
                        if (elemento.Category == null)
                        {
                            continue;
                        }

                        if (!elemento.Category.AllowsBoundParameters)
                        {
                            continue;
                        }

                        Parameter parametro =
                            elemento.LookupParameter(PARAMETRO_NOMBRE);

                        if (parametro == null)
                        {
                            CrearParametroParaCategoria(
                                doc,
                                uiApp.Application,
                                elemento.Category);

                            doc.Regenerate();
                        }
                    }

                    // ----------------------------------------------------
                    // SEGUNDO: ESCRIBIR LAS COORDENADAS
                    // ----------------------------------------------------

                    foreach (Element elemento in elementosSeleccionados)
                    {
                        // =================================================
                        // OBTENER PUNTO
                        // =================================================

                        XYZ puntoInsercion =
                            ObtenerPuntoInsercion(elemento);

                        if (puntoInsercion == null)
                        {
                            sinUbicacion++;
                            continue;
                        }

                        // =================================================
                        // OBTENER POSICION DEL PROYECTO
                        // =================================================

                        ProjectLocation ubicacion =
                            doc.ActiveProjectLocation;

                        if (ubicacion == null)
                        {
                            sinUbicacion++;
                            continue;
                        }

                        ProjectPosition posicion =
                            ubicacion.GetProjectPosition(
                                puntoInsercion);

                        if (posicion == null)
                        {
                            sinUbicacion++;
                            continue;
                        }

                        // =================================================
                        // CONVERTIR A METROS
                        // =================================================

                        double x =
                            UnitUtils.ConvertFromInternalUnits(
                                posicion.EastWest,
                                UnitTypeId.Meters);

                        double y =
                            UnitUtils.ConvertFromInternalUnits(
                                posicion.NorthSouth,
                                UnitTypeId.Meters);

                        double z =
                            UnitUtils.ConvertFromInternalUnits(
                                posicion.Elevation,
                                UnitTypeId.Meters);

                        // =================================================
                        // FORMATO
                        // =================================================

                        string coordenada =
                            FormatearCoordenada(
                                x,
                                y,
                                z);

                        // =================================================
                        // PARAMETRO
                        // =================================================

                        Parameter parametro =
                            elemento.LookupParameter(
                                PARAMETRO_NOMBRE);

                        if (parametro == null)
                        {
                            sinParametro++;
                            continue;
                        }

                        if (parametro.IsReadOnly)
                        {
                            sinParametro++;
                            continue;
                        }

                        if (parametro.StorageType !=
                            StorageType.String)
                        {
                            sinParametro++;
                            continue;
                        }

                        // =================================================
                        // ESCRIBIR
                        // =================================================

                        parametro.Set(coordenada);

                        procesados++;
                    }

                    trans.Commit();
                }

                // ========================================================
                // 4. RESULTADO
                // ========================================================

                string resultado =
                    "Proceso terminado.\n\n" +
                    "Elementos seleccionados: " +
                    elementosSeleccionados.Count +
                    "\n" +
                    "Elementos actualizados: " +
                    procesados;

                if (sinUbicacion > 0)
                {
                    resultado +=
                        "\nElementos sin ubicación: " +
                        sinUbicacion;
                }

                if (sinParametro > 0)
                {
                    resultado +=
                        "\nElementos sin parámetro actualizable: " +
                        sinParametro;
                }

                TaskDialog.Show(
                    "Coordenada XYZ",
                    resultado);

                return Result.Succeeded;
            }
            catch (
                Autodesk.Revit.Exceptions
                .OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.ToString();

                TaskDialog.Show(
                    "Error - Coordenada XYZ",
                    ex.Message);

                return Result.Failed;
            }
        }

        // ================================================================
        // OBTENER PUNTO DE INSERCION
        // ================================================================

        private XYZ ObtenerPuntoInsercion(Element elemento)
        {
            if (elemento == null)
            {
                return null;
            }

            XYZ puntoBase = null;

            Location ubicacion =
                elemento.Location;

            // ------------------------------------------------------------
            // LOCATION POINT
            // ------------------------------------------------------------

            LocationPoint ubicacionPunto =
                ubicacion as LocationPoint;

            if (ubicacionPunto != null)
            {
                puntoBase =
                    ubicacionPunto.Point;
            }

            // ------------------------------------------------------------
            // LOCATION CURVE
            // ------------------------------------------------------------

            if (puntoBase == null)
            {
                LocationCurve ubicacionCurva =
                    ubicacion as LocationCurve;

                if (ubicacionCurva != null)
                {
                    Curve curva =
                        ubicacionCurva.Curve;

                    if (curva != null)
                    {
                        puntoBase =
                            curva.GetEndPoint(0);
                    }
                }
            }

            // ------------------------------------------------------------
            // SI NO TENEMOS LOCATION, USAR BOUNDING BOX
            // ------------------------------------------------------------

            BoundingBoxXYZ caja =
                elemento.get_BoundingBox(null);

            if (puntoBase == null && caja != null)
            {
                XYZ centroLocal =
                    new XYZ(
                        (caja.Min.X + caja.Max.X) / 2.0,
                        (caja.Min.Y + caja.Max.Y) / 2.0,
                        (caja.Min.Z + caja.Max.Z) / 2.0);

                puntoBase =
                    caja.Transform.OfPoint(
                        centroLocal);
            }

            if (puntoBase == null)
            {
                return null;
            }

            // ------------------------------------------------------------
            // ELEVACION SUPERIOR
            //
            // Para elementos estructurales como pilotes, cimentaciones,
            // etc., utilizamos el parámetro nativo de Revit:
            //
            // STRUCTURAL_ELEVATION_AT_TOP
            //
            // Esto permite obtener la cota superior real del elemento
            // en lugar de utilizar el punto de inserción.
            // ------------------------------------------------------------

            Parameter parametroCotaSuperior =
                elemento.get_Parameter(
                    BuiltInParameter.STRUCTURAL_ELEVATION_AT_TOP);

            if (parametroCotaSuperior != null &&
                parametroCotaSuperior.StorageType ==
                StorageType.Double)
            {
                double zSuperior =
                    parametroCotaSuperior.AsDouble();

                return new XYZ(
                    puntoBase.X,
                    puntoBase.Y,
                    zSuperior);
            }

            // ------------------------------------------------------------
            // RESPALDO:
            // SI NO EXISTE "ELEVATION AT TOP", UTILIZAR EL MAXIMO
            // DE LA BOUNDING BOX.
            // ------------------------------------------------------------

            if (caja != null)
            {
                XYZ puntoSuperiorLocal =
                    new XYZ(
                        caja.Min.X,
                        caja.Min.Y,
                        caja.Max.Z);

                XYZ puntoSuperior =
                    caja.Transform.OfPoint(
                        puntoSuperiorLocal);

                return new XYZ(
                    puntoBase.X,
                    puntoBase.Y,
                    puntoSuperior.Z);
            }

            // ------------------------------------------------------------
            // ULTIMO RECURSO:
            // DEVOLVER EL PUNTO ORIGINAL
            // ------------------------------------------------------------

            return puntoBase;
        }

        // ================================================================
        // FORMATEAR COORDENADAS
        // ================================================================

        private string FormatearCoordenada(
            double x,
            double y,
            double z)
        {
            CultureInfo cultura =
                CultureInfo.InvariantCulture;

            return "("
                + x.ToString("0.000", cultura)
                + " , "
                + y.ToString("0.000", cultura)
                + " , "
                + z.ToString("0.000", cultura)
                + ")";
        }

        // ================================================================
        // CREAR / ACTUALIZAR PARAMETRO PARA UNA CATEGORIA
        // ================================================================

        private void CrearParametroParaCategoria(
            Document doc,
            Autodesk.Revit.ApplicationServices.Application app,
            Category categoria)
        {
            if (categoria == null)
            {
                return;
            }

            if (!categoria.AllowsBoundParameters)
            {
                return;
            }

            // ============================================================
            // BUSCAR PARAMETRO EXISTENTE
            // ============================================================

            BindingMap mapaVinculaciones =
                doc.ParameterBindings;

            DefinitionBindingMapIterator iterador =
                mapaVinculaciones.ForwardIterator();

            Definition definicionExistente = null;

            ElementBinding vinculacionExistente = null;

            CategorySet categoriasExistentes =
                app.Create.NewCategorySet();

            iterador.Reset();

            while (iterador.MoveNext())
            {
                Definition definicionActual =
                    iterador.Key;

                if (definicionActual == null)
                {
                    continue;
                }

                if (!definicionActual.Name.Equals(
                    PARAMETRO_NOMBRE,
                    StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                definicionExistente =
                    definicionActual;

                // Current devuelve object en Revit 2024.
                Binding vinculacionActual =
                    iterador.Current as Binding;

                if (vinculacionActual == null)
                {
                    throw new InvalidOperationException(
                        "No se pudo obtener la vinculación existente " +
                        "del parámetro '" +
                        PARAMETRO_NOMBRE +
                        "'.");
                }

                vinculacionExistente =
                    vinculacionActual as ElementBinding;

                if (vinculacionExistente == null)
                {
                    throw new InvalidOperationException(
                        "El parámetro '" +
                        PARAMETRO_NOMBRE +
                        "' no está vinculado a elementos.");
                }

                // --------------------------------------------------------
                // CONSERVAR CATEGORIAS EXISTENTES
                // --------------------------------------------------------

                foreach (Category categoriaExistente
                    in vinculacionExistente.Categories)
                {
                    AgregarCategoriaSiFalta(
                        categoriasExistentes,
                        categoriaExistente);
                }

                break;
            }

            // ============================================================
            // AGREGAR CATEGORIA ACTUAL
            // ============================================================

            AgregarCategoriaSiFalta(
                categoriasExistentes,
                categoria);

            // ============================================================
            // CREAR DEFINICION SI NO EXISTE
            // ============================================================

            if (definicionExistente == null)
            {
                definicionExistente =
                    ObtenerOCrearDefinicionCompartida(
                        app);
            }

            // ============================================================
            // CREAR BINDING DE INSTANCIA
            // ============================================================

            InstanceBinding nuevaVinculacion =
                app.Create.NewInstanceBinding(
                    categoriasExistentes);

            // ============================================================
            // INSERTAR / REINSERTAR
            // ============================================================

            bool resultado;

            if (vinculacionExistente == null)
            {
                resultado =
                    mapaVinculaciones.Insert(
                        definicionExistente,
                        nuevaVinculacion,
                        GroupTypeId.Data);
            }
            else
            {
                resultado =
                    mapaVinculaciones.ReInsert(
                        definicionExistente,
                        nuevaVinculacion,
                        GroupTypeId.Data);
            }

            if (!resultado)
            {
                throw new InvalidOperationException(
                    "Revit no pudo crear o actualizar el parámetro '" +
                    PARAMETRO_NOMBRE +
                    "' para la categoría '" +
                    categoria.Name +
                    "'.");
            }
        }

        // ================================================================
        // AGREGAR CATEGORIA SIN DUPLICAR
        // ================================================================

        private void AgregarCategoriaSiFalta(
            CategorySet conjunto,
            Category categoria)
        {
            if (conjunto == null ||
                categoria == null)
            {
                return;
            }

            foreach (Category existente in conjunto)
            {
                if (existente.Id.Equals(
                    categoria.Id))
                {
                    return;
                }
            }

            conjunto.Insert(categoria);
        }

        // ================================================================
        // OBTENER / CREAR DEFINICION COMPARTIDA
        // ================================================================

        private Definition ObtenerOCrearDefinicionCompartida(
            Autodesk.Revit.ApplicationServices.Application app)
        {
            string carpetaTemporal =
                Path.Combine(
                    Path.GetTempPath(),
                    "ICCU_Revit");

            if (!Directory.Exists(carpetaTemporal))
            {
                Directory.CreateDirectory(
                    carpetaTemporal);
            }

            string archivoParametros =
                Path.Combine(
                    carpetaTemporal,
                    "ICCU_SharedParameters.txt");

            // ============================================================
            // CREAR ARCHIVO SI NO EXISTE
            // ============================================================

            if (!File.Exists(
                archivoParametros))
            {
                string contenidoArchivo =
                    "# This is a Revit shared parameter file.\r\n" +
                    "# Do not edit manually.\r\n" +
                    "*META\tVERSION\tMINVERSION\r\n" +
                    "META\t2\t1\r\n" +
                    "*GROUP\tID\tNAME\r\n" +
                    "*PARAM\tGUID\tNAME\tDATATYPE\tDATACATEGORY\t" +
                    "GROUP\tVISIBLE\tDESCRIPTION\tUSERMODIFIABLE\r\n";

                File.WriteAllText(
                    archivoParametros,
                    contenidoArchivo,
                    new UTF8Encoding(false));
            }

            // ============================================================
            // GUARDAR ARCHIVO ORIGINAL
            // ============================================================

            string archivoOriginal =
                app.SharedParametersFilename;

            try
            {
                app.SharedParametersFilename =
                    archivoParametros;

                DefinitionFile archivoDefiniciones =
                    app.OpenSharedParameterFile();

                if (archivoDefiniciones == null)
                {
                    throw new InvalidOperationException(
                        "No fue posible abrir el archivo de parámetros " +
                        "compartidos.");
                }

                // ========================================================
                // GRUPO ICCU
                // ========================================================

                DefinitionGroup grupo =
                    archivoDefiniciones.Groups.get_Item(
                        GRUPO_PARAMETRO);

                if (grupo == null)
                {
                    grupo =
                        archivoDefiniciones.Groups.Create(
                            GRUPO_PARAMETRO);
                }

                // ========================================================
                // BUSCAR PARAMETRO
                // ========================================================

                foreach (Definition definicion
                    in grupo.Definitions)
                {
                    if (definicion.Name.Equals(
                        PARAMETRO_NOMBRE,
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return definicion;
                    }
                }

                // ========================================================
                // CREAR PARAMETRO DE TEXTO
                // ========================================================

                ExternalDefinitionCreationOptions opciones =
                    new ExternalDefinitionCreationOptions(
                        PARAMETRO_NOMBRE,
                        SpecTypeId.String.Text);

                opciones.Visible = true;

                opciones.UserModifiable = true;

                opciones.Description =
                    "Coordenadas X, Y y Z del elemento.";

                Definition nuevaDefinicion =
                    grupo.Definitions.Create(
                        opciones);

                if (nuevaDefinicion == null)
                {
                    throw new InvalidOperationException(
                        "No fue posible crear el parámetro '" +
                        PARAMETRO_NOMBRE +
                        "'.");
                }

                return nuevaDefinicion;
            }
            finally
            {
                // ========================================================
                // RESTAURAR CONFIGURACION ORIGINAL
                // ========================================================

                app.SharedParametersFilename =
                    archivoOriginal;
            }
        }
    }
}