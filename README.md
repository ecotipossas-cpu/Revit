# Revit Tools

Conjunto de herramientas de automatización desarrolladas en C# para Autodesk Revit.

El repositorio está orientado al desarrollo de comandos para automatizar procesos BIM, modelado, gestión de información y preparación de modelos.

---

## Tools

### Crear nuevos parametros

Comando para automatizar la creación y asignación de parámetros compartidos en Autodesk Revit.

La herramienta utiliza una matriz Excel como fuente de información.

#### Funcionalidades

- Selección del archivo Excel desde el comando.
- Selección de la hoja de trabajo mediante número.
- Lectura de la matriz de parámetros.
- Creación automática de Shared Parameters.
- Asignación de parámetros a múltiples categorías.
- Configuración de parámetros como:
  - Instance
  - Type
- Configuración del grupo de parámetros de Revit.
- Configuración del tipo de parámetro.
- Creación y actualización de bindings.
- Detección de parámetros existentes para evitar duplicados.
- Reporte del resultado del proceso.

#### Matriz Excel

La matriz debe contener las siguientes columnas:

| Columna | Descripción |
|---|---|
| Grupo | Grupo de parámetros dentro de Revit |
| Categorias | Categorías a las que se asignará el parámetro |
| Parametros | Nombre del parámetro |
| Tipo | Tipo de dato del parámetro |
| Binding | Instance o Type |

Las categorías múltiples deben estar separadas por comas.

Ejemplo:

| Grupo | Categorias | Parametros | Tipo | Binding |
|---|---|---|---|---|
| Construction | Walls,Railings,Generic Models | ICCU_ElementKey | Text | Instance |
| Construction | Generic Models | BIM_Quantities_Volumen | Volume | Instance |

---

## Requisitos

- Autodesk Revit 2024
- Visual Studio
- .NET Framework 4.8
- C#
- ClosedXML
- Revit API

---

## Estructura

```text
Revit/
│
├── .gitignore
├── README.md
│
└── Crear nuevos parametros/
    ├── Creacion de parametros.sln
    │
    └── Creacion de parametros/
        ├── Command.cs
        ├── Creacion de parametros.csproj
        ├── packages.config
        └── Properties/