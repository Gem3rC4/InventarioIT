# Conectar PowerShell con el Excel de SharePoint

Este proyecto envía los 18 campos del inventario directamente al archivo de Excel mediante Microsoft Graph. No requiere un flujo de Power Automate ni su licencia Premium. Sí necesita una cuenta de Microsoft 365 con acceso de escritura al archivo de SharePoint o OneDrive para la Empresa.

La persona que ejecuta el comando inicia sesión con Microsoft cuando se solicita. La integración actúa con el acceso de esa persona: no puede escribir en un archivo que su cuenta no pueda editar. GitHub almacena el código; los datos se guardan en el Excel y en respaldos locales.

## 1. Preparar la tabla existente

1. Abra el archivo `.xlsx` y confirme que puede editarlo con la cuenta que usará desde PowerShell.
2. Confirme que las 18 columnas están dentro de una **tabla de Excel**. Si solamente son encabezados en un rango, seleccione el rango y use **Insertar > Tabla**, indicando que tiene encabezados.
3. Consulte el nombre en **Diseño de tabla > Nombre de la tabla**. Puede usar `InventarioEquipos`.
4. Copie el enlace completo desde **Compartir > Copiar vínculo**. El enlace recortado visible en una captura de pantalla no basta para localizar el archivo.

Los encabezados esperados son exactamente estos:

```text
Equipo | Usuario | Marca | Modelo | Serie | Procesador | TipoEquipo | RAM_GB | DiscoTotal_GB | DiscoUsado_GB | DiscoLibre_GB | TipoDisco | IP | Windows | Version | Activado | Fecha | Area
```

El cliente verifica los encabezados antes de escribir y acomoda los valores al orden de las columnas de la tabla. Si el libro tiene una sola tabla, puede seleccionarla automáticamente; si hay varias, indique su nombre para evitar escribir en otra tabla.

Configure `Serie`, `IP` y `Version` como **Texto** si necesita conservar ceros iniciales. `Fecha` usa la hora local del equipo con formato `yyyy-MM-dd HH:mm:ss`; puede usar Texto también para esa columna. Las cuatro columnas terminadas en `_GB` reciben números o celdas vacías si el dato no pudo consultarse.

## 2. Preparar Windows PowerShell 5.1

Microsoft Graph PowerShell admite Windows PowerShell 5.1 con **.NET Framework 4.7.2 o posterior** y PowerShellGet actualizado. PowerShell 7 también está admitido. Instale el módulo desde la misma versión de PowerShell donde ejecutará el inventario.

Solo se necesita el módulo de autenticación:

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery
```

El primer comando habilita TLS 1.2 en ese proceso para descargar y usar servicios HTTPS. La instalación puede pedir autorización para instalar el proveedor NuGet o usar PowerShell Gallery. `CurrentUser` instala el módulo para su usuario sin requerir una instalación global como administrador. También puede ejecutar el inventario con `-InstalarDependencias` para solicitar la instalación si falta el módulo.

Puede comprobarlo con:

```powershell
Get-Module -ListAvailable Microsoft.Graph.Authentication | Select-Object Name, Version
```

No es necesario instalar todos los módulos de `Microsoft.Graph`, Excel de escritorio ni complementos para manipular el archivo.

## 3. Configurar el destino y comprobarlo

Introduzca el enlace completo del libro y compruebe la conexión sin insertar una fila:

```powershell
$enlaceInventario = Read-Host 'Pegue el enlace completo del Excel'
.\Inventario-Equipo.ps1 -ExcelUrl $enlaceInventario -Tabla 'InventarioEquipos' -ComprobarDestino
```

Cambie el nombre de la tabla del ejemplo por el de su archivo. Si el libro tiene una sola tabla, puede omitir `-Tabla`. No incluya tokens ni credenciales en la dirección.

El inicio de sesión usa el flujo de código de dispositivo:

1. PowerShell muestra la dirección de Microsoft y un código temporal.
2. Abra la dirección que muestra el comando, escriba el código e inicie sesión con la cuenta autorizada para editar el Excel. Use el perfil del navegador de esa cuenta para evitar seleccionar otra sesión de Microsoft 365.
3. Si aparece una pantalla de consentimiento, revise la aplicación y los permisos solicitados. Regrese a PowerShell después de completar el inicio de sesión.

El permiso solicitado para leer la estructura y agregar una fila es **`Files.ReadWrite`**, de tipo delegado. Microsoft Graph PowerShell gestiona el token de Microsoft; no se usa un token de GitHub para editar Excel. El cliente usa la aplicación de Microsoft Graph PowerShell y no necesita que cree una aplicación propia para el caso normal.

La sesión se limita al proceso mediante `-ContextScope Process`. Cada nueva sesión de PowerShell puede requerir autenticación de nuevo. No guarde contraseñas, tokens de Microsoft o credenciales de GitHub dentro del archivo de configuración.

Para repetir el destino sin escribirlo cada vez, copie el archivo de ejemplo de configuración siguiendo el README. Sus claves son `ExcelUrl`, `Tabla` y, opcionalmente, `TenantId` y `ClientId`. Los dos últimos permiten que un administrador especifique un tenant o una aplicación propia cuando la organización lo necesite. El archivo real de configuración local está excluido del repositorio; no publique enlaces privados del Excel.

## 4. Recopilar y enviar

Primero puede comprobar la recopilación sin enviar datos:

```powershell
.\Inventario-Equipo.ps1 -Area 'Sistemas' -SoloLocal
```

Para registrar el equipo en Excel:

```powershell
.\Inventario-Equipo.ps1 -Area 'Sistemas' -ExcelUrl $enlaceInventario -Tabla 'InventarioEquipos'
```

Si omite `-Area`, el comando solicita el departamento antes de continuar. Haga primero una prueba con un equipo y compruebe que aparezca una sola fila con sus 18 campos dentro de la tabla correcta. La conectividad de producción queda verificada únicamente después de esa prueba con la cuenta y el archivo reales.

El script conserva una copia CSV y una copia JSON antes del envío. Si hay un error de permisos o conexión, esos archivos permiten recuperar el registro. El mensaje del cliente indica sus rutas.

Cada ejecución crea un registro nuevo: no reemplaza automáticamente el inventario anterior del mismo equipo. Si hubo un tiempo de espera o un error de red, revise el Excel por `Equipo` y `Fecha` antes de volver a ejecutar o reenviar el respaldo. La escritura pudo haberse completado aunque el cliente no recibiera confirmación.

Si confirmó que el registro no está en el Excel, puede enviar el JSON original mediante `-RegistroJson`, conservando la fecha original:

```powershell
$rutaRespaldo = Read-Host 'Pegue la ruta completa del respaldo JSON que revisó'
.\Inventario-Equipo.ps1 -RegistroJson $rutaRespaldo -ExcelUrl $enlaceInventario -Tabla 'InventarioEquipos'
```

La inserción se realiza sin reintentos automáticos para reducir el riesgo de duplicados. La tabla no impone unicidad y la API de agregar filas no ofrece una garantía de una sola inserción por solicitud. No se debe prometer que los duplicados son imposibles.

## Permisos y políticas de la organización

Si aparece **Se necesita la aprobación del administrador**, solicite al administrador de Microsoft 365 que revise el consentimiento para Microsoft Graph PowerShell y el permiso delegado `Files.ReadWrite`. Aunque el permiso no requiere siempre consentimiento de administrador por definición, el tenant puede restringir el consentimiento de sus usuarios o el uso de aplicaciones.

Las políticas de acceso condicional también pueden impedir el código de dispositivo o exigir condiciones adicionales. En ese caso se necesita que el administrador revise el método de acceso permitido. No sirve reemplazar este permiso por una clave secreta y autenticación de aplicación: la operación de Graph para agregar filas a Excel **no admite permisos de aplicación**.

El enlace identifica el archivo; no sustituye los permisos de edición ni garantiza que la cuenta autenticada tenga acceso. Si ve `403`, confirme cuenta, consentimiento y acceso al archivo. Si ve `404`, confirme el vínculo completo, el acceso de esa cuenta y que el archivo no se haya movido o eliminado. Si el vínculo no puede resolverse, solicite uno nuevo al propietario; una captura o una dirección recortada no sirve.

## Uso en varios equipos

Microsoft recomienda enviar las escrituras a cada libro de forma secuencial. Ejecutar varios equipos al mismo tiempo contra el mismo archivo puede causar limitación de solicitudes, conflictos o tiempos de espera. Este cliente no coordina entre diferentes equipos: para un lote moderado, ejecute uno después del otro.

Si requiere ejecución desatendida, cientos de equipos simultáneos o una restricción estricta contra duplicados, hará falta una recepción central con cola y almacenamiento de solicitudes; no basta con distribuir este comando y una contraseña. La integración aquí incluida está diseñada para un inventario iniciado por una persona con acceso al Excel.

## Fuentes oficiales

- [Instalación de Microsoft Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation): compatibilidad con PowerShell 5.1, .NET Framework y módulos individuales.
- [Comandos de autenticación](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands): inicio de sesión por dispositivo, permisos delegados, aplicación propia opcional y alcance de la sesión.
- [Acceso a archivos compartidos](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0): resolución de vínculos mediante Graph.
- [Agregar filas a una tabla de Excel](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-add?view=graph-rest-1.0): permiso `Files.ReadWrite`, cuerpo de la solicitud y ausencia de permisos de aplicación.
- [Buenas prácticas para Excel en Graph](https://learn.microsoft.com/en-us/graph/workbook-best-practice): sesiones, escrituras secuenciales y limitación de solicitudes.
- [Contexto de solicitudes de Graph PowerShell](https://learn.microsoft.com/en-us/powershell/module/microsoft.graph.authentication/set-mgrequestcontext?view=graph-powershell-1.0): configuración de reintentos.
- [Consentimiento de aplicaciones en Microsoft Entra](https://learn.microsoft.com/en-us/entra/identity/enterprise-apps/configure-user-consent): políticas que puede fijar el administrador.
