# Lector de registros de conexión (Proyecto de consola - VB.NET)

Este proyecto es una aplicación de consola en Visual Basic .NET que carga un archivo Excel con registros de conexión (Nombre, Día, Hora de conexión, Hora de desconexión) y los muestra en la terminal. Está organizado usando principios básicos de POO y encapsulamiento.

**Archivos principales**
- `Conexion.vb`: clase que representa una fila del Excel (atributos privados y propiedades públicas).
- `ExcelLoader.vb`: clase responsable de leer el archivo Excel y devolver una lista de `Conexion`.
- `App.vb`: lógica principal de la aplicación (pide ruta, muestra tabla, maneja reintentos).
- `Program.vb`: punto de entrada (`Main`) que llama a `App.Run()`.

**Requisitos**
- .NET SDK instalado (se probó con .NET 9).
- Paquete NuGet: `EPPlus` para leer archivos Excel.

Si no tienes el SDK, instálalo desde https://dotnet.microsoft.com/

**Instalación de dependencias (PowerShell)**
Ejecuta en la carpeta del proyecto `lector`:

```powershell
cd "c:\Users\Pepito_Windows\Downloads\Universidad 2025-2\Programación\10% 2do corte\lector"
dotnet add .\lector.vbproj package EPPlus --version 6.3.2
dotnet restore
```

Nota: NuGet puede resolver una versión más reciente de EPPlus; eso está bien (aparecerá una advertencia NU1603 si se usa otra versión).

**Compilar y ejecutar**

```powershell
dotnet build .\lector.vbproj
dotnet run --project .\lector.vbproj
```

La aplicación es interactiva: al ejecutarla verás el prompt
`Ruta del archivo Excel (o Enter para salir):` — escribe la ruta completa del archivo `.xlsx` y presiona Enter.

**Formato del archivo Excel**
La hoja debe tener en la primera fila los encabezados (pueden tener variantes):

| Nombre | Día | Hora de conexión | Hora de desconexión |
|--------|-----|------------------|---------------------|

Ejemplos de encabezados aceptados (no son sensibles a mayúsculas ni acentos):
- Nombre / Nombre completo / Name
- Día / Fecha / Date
- Hora de conexión / Hora entrada / Entrada / TimeIn
- Hora de desconexión / Hora salida / Salida / TimeOut

El programa lee desde la segunda fila hacia abajo y crea objetos `Conexion` por fila.

**Salida esperada**
Se imprimirá una tabla con columnas: `Nombre`, `Día`, `Entrada`, `Salida`.

**Notas importantes**
- EPPlus usa una licencia comercial; en este proyecto se configura `LicenseContext = NonCommercial` porque es un trabajo académico. Si planeas uso comercial, revisa la licencia de EPPlus.
- Si prefieres usar `Microsoft.Office.Interop.Excel` (requiere Excel instalado y referencias COM), indícalo y te doy instrucciones.

**Sugerencia rápida para crear un archivo de ejemplo**
1. Abre Excel y crea una hoja con los encabezados mencionados.
2. Guarda como `sample.xlsx` en cualquier carpeta.
3. Ejecuta la app y proporciona la ruta completa, por ejemplo: `C:\Users\TuUsuario\Desktop\sample.xlsx`.

Si quieres, puedo crear un `sample.xlsx` de ejemplo automáticamente en el proyecto y ejecutar la app contra él. Dime si lo hago.

---
Proyecto preparado como ejercicio de programación orientada a objetos (encapsulamiento y separación de responsabilidades).
