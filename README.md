# InventarioIT

Un comando para Windows PowerShell 5.1 que recopila el inventario del equipo, guarda respaldos locales y agrega una fila a la tabla de un archivo Excel de SharePoint/OneDrive para la Empresa. Usa Microsoft Graph con inicio de sesión de Microsoft 365. No requiere Power Automate ni Excel de escritorio.

El Excel sigue siendo el destino de los datos; este repositorio almacena únicamente el código. No incluye vínculos privados, contraseñas ni tokens.

## Ejecutar desde PowerShell

Abra **Windows PowerShell** con su usuario habitual. Este comando descarga el código de este repositorio y lo ejecuta. Si el módulo de autenticación falta, solicita instalarlo para su usuario desde PowerShell Gallery:

```powershell
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$codigoInventario = Invoke-RestMethod -Uri 'https://raw.githubusercontent.com/Gem3rC4/InventarioIT/main/Inventario-Equipo.ps1'
& ([scriptblock]::Create($codigoInventario)) -InstalarDependencias
```

Le pedirá el **área**, el **enlace completo del Excel** y que inicie sesión con Microsoft. Use una cuenta que pueda editar ese archivo. Si el libro contiene una sola tabla, la detecta; si contiene varias, agregue `-Tabla 'NombreReal'` al comando. Para una versión fija del código, sustituya `main` en la dirección por el identificador del commit revisado.

La primera instalación puede preguntar por NuGet y PowerShell Gallery. Windows PowerShell 5.1 necesita .NET Framework 4.7.2 o posterior. Si la organización solicita aprobación de administrador al iniciar sesión, el administrador debe revisar el consentimiento de Microsoft Graph PowerShell para el permiso delegado `Files.ReadWrite`. El enlace del Excel no reemplaza ese permiso.

También puede descargar el repositorio y ejecutar el archivo local:

```powershell
.\Inventario-Equipo.ps1 -InstalarDependencias
```

Si Windows bloquea un archivo descargado, después de revisar su contenido puede desbloquear ese archivo con `Unblock-File .\Inventario-Equipo.ps1` o ejecutar una sola vez con `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Inventario-Equipo.ps1 -InstalarDependencias`. No hace falta cambiar permanentemente la política del equipo.

## Guardar el destino en este equipo

Para no volver a pegar el enlace:

```powershell
$carpetaInventario = Join-Path $env:LOCALAPPDATA 'InventarioEquipos'
$null = New-Item -ItemType Directory -Path $carpetaInventario -Force
@{
    ExcelUrl = (Read-Host 'Pegue el enlace completo del Excel')
    Tabla = ''
    TenantId = ''
    ClientId = ''
} | ConvertTo-Json | Set-Content -LiteralPath (Join-Path $carpetaInventario 'config.json') -Encoding UTF8
```

La tabla puede quedar vacía si solo hay una. El script busca esa configuración automáticamente. Puede usar otra ubicación con `-ConfigPath 'C:\Ruta\config.local.json'`; los parámetros explícitos tienen prioridad. `config.example.json` contiene las cuatro claves de ejemplo. Los campos `TenantId` y `ClientId` son opcionales, para una configuración corporativa indicada por el administrador. Una aplicación propia debe tener `Files.ReadWrite` delegado y estar habilitada como cliente público para autenticación por dispositivo; no usa secreto de aplicación.

## Comprobar antes del primer envío

```powershell
# Con configuración guardada: valida archivo, tabla y encabezados, sin insertar.
.\Inventario-Equipo.ps1 -ComprobarDestino

# Recopila sin conectarse a Microsoft ni enviar a Excel.
.\Inventario-Equipo.ps1 -Area 'Sistemas' -SoloLocal

# Agrega una fila de inventario; solicita iniciar sesión.
.\Inventario-Equipo.ps1 -Area 'Sistemas'
```

El script comprueba los 18 nombres de columna exactos y adapta los valores a su orden real. Si faltan o sobran columnas, no inserta nada. Protege los textos que Excel podría interpretar como fórmulas.

| Campo | Dato |
| --- | --- |
| Equipo | Nombre del equipo |
| Usuario | Usuario de la sesión interactiva, sin dominio |
| Marca / Modelo / Serie | Fabricante, modelo y serie del BIOS |
| Procesador / TipoEquipo | Procesadores y clasificación del chasis |
| RAM_GB | Memoria física, en GiB redondeados a dos decimales |
| DiscoTotal_GB / DiscoUsado_GB / DiscoLibre_GB | Suma de volúmenes locales fijos, en GiB |
| TipoDisco | SSD/HDD detectados, sin USB/SD/MMC |
| IP | Direcciones IPv4 de adaptadores habilitados, salvo loopback y APIPA |
| Windows / Version / Activado | Edición, versión y estado de activación; no recopila claves de producto |
| Fecha / Area | Hora local del escaneo y departamento indicado |

Los nombres `_GB` se conservan para coincidir con la tabla existente; el cálculo usa 1 GiB = 1.073.741.824 bytes, igual que el script original. La suma de volúmenes no representa necesariamente la capacidad física instalada. Los datos que el hardware no informa aparecen como `Desconocido` o vacío en columnas numéricas. No se recopilan documentos, contraseñas, historial ni todos los programas instalados: se recopilan los 18 campos de esa tabla.

## Respaldos y reenvío

Antes de iniciar sesión, cada escaneo se guarda en `%LOCALAPPDATA%\InventarioEquipos` como JSON y CSV UTF-8. Puede cambiarlo con `-DirectorioSalida`. El JSON conserva los valores originales; el CSV neutraliza texto que pudiera interpretarse como fórmula al abrirlo.

Cada ejecución agrega **una fila nueva**. No actualiza automáticamente filas anteriores del mismo equipo. No hay reintentos automáticos: si la red falla después de escribir, podría existir una fila aunque el programa muestre error. Compruebe `Equipo` y `Fecha` en Excel antes de reenviar:

```powershell
.\Inventario-Equipo.ps1 -RegistroJson 'C:\Ruta\respaldo.json'
```

El reenvío conserva la fecha original y crea una nueva copia local. Ejecute el inventario de varios equipos **uno después del otro**; este cliente no mantiene una cola central ni coordina escrituras concurrentes al mismo libro.

## Validación del proyecto

Las pruebas de `tests/Probar-Inventario.ps1` usan datos sintéticos y respuestas simuladas de Graph. Se ejecutan en Windows PowerShell 5.1 sin iniciar sesión ni escribir en un Excel real:

```powershell
powershell.exe -NoProfile -File .\tests\Probar-Inventario.ps1
```

En el entorno restringido de desarrollo no se pudo consultar CIM del equipo real. La consulta de hardware y la inserción real requieren comprobarse desde una sesión normal de Windows, con la cuenta y el archivo de la organización. Una prueba simulada no confirma los permisos ni la conexión de producción.

Consulte [la guía de conexión y permisos](docs/Configurar-Excel.md) para la configuración detallada y las fuentes oficiales de Microsoft.
