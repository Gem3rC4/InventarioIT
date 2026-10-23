# InventarioIT

Un comando de Windows PowerShell 5.1 recopila los 18 campos del equipo, guarda respaldos y los envía a un **equipo Windows central**. Los equipos inventariados **no inician sesión en Microsoft**, no necesitan Excel instalado ni módulos de Graph.

Solo el central se conecta a Microsoft 365 para escribir en la tabla del Excel existente. No requiere Power Automate. Esta versión reemplaza el envío directo con autenticación en cada computadora.

## Preparación en el equipo central

Haga esta preparación en el Windows que permanecerá encendido, en la misma red que las demás computadoras:

1. [Descargue el proyecto](https://github.com/Gem3rC4/InventarioIT/archive/refs/heads/main.zip), extraiga el ZIP y abra **Windows PowerShell como administrador** en la carpeta del proyecto.
2. Ejecute:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Configurar-Servidor.ps1 -PrepararRed
```

El configurador detecta la IP privada del central y solicita el enlace completo del Excel. Si detecta varias direcciones, le pide elegir una. También puede indicar `-DireccionIP '192.168.1.10'`. Si el libro tiene varias tablas, agregue `-Tabla 'NombreReal'`; con una sola tabla, se detecta automáticamente.

`-PrepararRed` crea una reserva HTTP para el usuario del receptor y una regla de firewall limitada a la IP y puerto elegidos, la subred local y los perfiles Dominio/Privado. No cambia el perfil de red. Si la red de Windows del central figura como **Pública**, la regla no se aplica y los demás equipos reciben *"Se excedió el tiempo de espera"*; el configurador lo avisa. En ese caso agregue `-PermitirRedPublica`, que incluye también el perfil Público y mantiene la regla limitada a la IP, el puerto y la subred local. Sin `-PrepararRed` se generan únicamente los archivos locales.

La reserva se comprueba por el prefijo y el identificador de seguridad (SID) del usuario, que `netsh` no traduce, así que funciona con Windows en español o inglés. Si no existe, se crea para ese SID; si ya pertenece al usuario, se conserva. Si pertenece a otra cuenta, no se modifica y el mensaje indica cómo usar otro puerto (`-Puerto`) o eliminar esa reserva.

3. Inicie el receptor:

```powershell
& "$env:LOCALAPPDATA\InventarioEquipos\Iniciar-Servidor.cmd"
```

No se instala ningún módulo. El receptor inicia sesión en Microsoft directamente, con código de dispositivo, sin el módulo `Microsoft.Graph.Authentication`, que en Windows PowerShell 5.1 falla con *"Error al escribir en un agente de escucha"*. La consola muestra la dirección y el código, abre el navegador y copia el código al portapapeles.

La sesión se guarda en `%LOCALAPPDATA%\InventarioEquipos\sesion-microsoft.dat`, cifrada con DPAPI para esa cuenta de Windows. Al reiniciar el receptor se reutiliza y no pide un código nuevo mientras Microsoft la mantenga válida. Para iniciar con otra cuenta, agregue `-NuevaSesion` al iniciar el receptor o elimine ese archivo.

**Solo en el central**, siga el código temporal de Microsoft e inicie sesión con una cuenta que pueda editar el Excel. Si aparece que se necesita aprobación del administrador de Microsoft 365, este debe revisar `Files.ReadWrite` delegado para Microsoft Graph PowerShell. El vínculo anónimo de Excel no reemplaza la autorización de Graph.

Espere a que aparezca **Receptor activo**. Mantenga encendidos el central y su sesión de Windows, con esa ventana abierta y acceso a Internet. Use una IP fija o una reserva DHCP para conservar el comando generado. Al detener y reiniciar el receptor, Microsoft puede volver a solicitar autenticación; durante su funcionamiento renueva los tokens cuando la política de la organización lo permite.

La configuración real y la cola se guardan fuera del repositorio, en `%LOCALAPPDATA%\InventarioEquipos`. Ejecute el receptor con la misma cuenta de Windows que hizo la configuración de red. No se ha configurado como un servicio de Windows ni como una tarea de inicio automático.

## Comando para las demás computadoras

En el central, abra:

```powershell
Get-Content "$env:LOCALAPPDATA\InventarioEquipos\Comando-Equipos.txt" -Raw
```

**Copie ese comando y ejecútelo en PowerShell en cada computadora.** Ya contiene la dirección del central y su clave de recepción; solo pedirá el área o departamento. No abre una sesión de Microsoft en los clientes.

El comando contiene una clave de recepción: compártalo con el personal que realizará el inventario y no lo publique en GitHub. La clave permite enviar solicitudes al receptor; no es una contraseña de Microsoft.

También puede copiar `cliente.config.json`, generado en esa misma carpeta, a `%LOCALAPPDATA%\InventarioEquipos\cliente.config.json` de un cliente y ejecutar el script descargado:

```powershell
.\Inventario-Equipo.ps1
```

El archivo `config.example.json` muestra las claves `ServidorUrl` y `ClaveApi`. Los parámetros explícitos tienen prioridad sobre la configuración local.

## Resultado del envío

- **RegistradoEnExcel:** el central recibió confirmación de Excel.
- **EnCola:** el central guardó la solicitud en disco y la procesará. Todavía no confirma una fila en Excel.
- **Requiere revisión:** la inserción falló o su resultado es incierto; revise el central y Excel antes de cualquier reenvío.

El cliente espera unos segundos la confirmación. Con `-EsperarSegundos 0` devuelve la recepción inicial; la cola sigue procesándose aunque el cliente cierre PowerShell. El receptor escribe una fila a la vez. Una escritura lenta puede retrasar otras respuestas; los clientes conservan su solicitud para reintentar con el mismo identificador.

## Campos y tabla de Excel

La tabla debe contener exactamente estos 18 encabezados:

```text
Equipo | Usuario | Marca | Modelo | Serie | Procesador | TipoEquipo | RAM_GB | DiscoTotal_GB | DiscoUsado_GB | DiscoLibre_GB | TipoDisco | IP | Windows | Version | Activado | Fecha | Area
```

El receptor vuelve a comprobar los encabezados antes de cada inserción y coloca los valores según su orden real. Si faltan o sobran columnas, conserva el registro para revisión. Cada captura agrega una fila nueva; no reemplaza inventarios anteriores del mismo equipo.

Se recopilan nombre del equipo, usuario interactivo sin dominio, fabricante, modelo, serie, procesadores, tipo de chasis, RAM, volúmenes fijos y su espacio, tipo SSD/HDD, IPv4, edición y versión de Windows, activación, fecha local y área. No recopila documentos, contraseñas ni claves de producto. Los nombres `_GB` se conservan para coincidir con el Excel; el cálculo usa GiB, igual que el script original. La capacidad suma volúmenes fijos y no necesariamente equivale a capacidad física instalada.

## Respaldos y recuperación

Antes de enviar, el cliente guarda CSV, JSON y `*.solicitud.json` en `%LOCALAPPDATA%\InventarioEquipos`. El último conserva el identificador único y el contenido de esa captura. Para recopilar sin enviar:

```powershell
.\Inventario-Equipo.ps1 -Area 'Sistemas' -SoloLocal
```

Si no recibió respuesta del central, reenvíe **la misma solicitud**, conservando su identificador:

```powershell
.\Inventario-Equipo.ps1 -SolicitudJson 'C:\Ruta\equipo.solicitud.json'
```

Si ese ID ya existe con el mismo contenido, el receptor devuelve su estado sin agregar otra solicitud. Si el contenido cambió, lo rechaza. No cree una captura nueva como sustituto de un reintento.

El receptor persiste la cola antes de responder y conserva los recibos enviados. Si se detuvo durante una inserción, marca esa solicitud `needs-review` al reiniciar. No la inserta de nuevo automáticamente. Con los 18 campos actuales no puede garantizar exactamente una inserción si Excel escribió pero la confirmación se perdió. Compare Equipo y Fecha antes de resolver estos casos; no borre los recibos de la cola para intentar forzar un envío.

## Alcance de la red

La configuración generada usa HTTP **solo en una red local privada y confiable**. En esa modalidad el inventario y la clave viajan sin cifrado. No publique el puerto en Internet. Para una red no confiable, use HTTPS con un certificado válido; el receptor admite un prefijo HTTPS privado configurado en HTTP.sys. No desactiva la validación de certificados del cliente.

La clave, el enlace del Excel, la configuración real y los registros permanecen fuera del repositorio. `servidor.config.example.json` contiene únicamente valores ilustrativos. La cola admite hasta 1000 solicitudes pendientes; conserve suficiente espacio en disco y revise los estados que requieren intervención.

## Pruebas

Las suites usan datos sintéticos y respuestas simuladas, sin iniciar sesión ni modificar un Excel real:

```powershell
powershell.exe -NoProfile -File .\tests\Probar-Cliente.ps1
powershell.exe -NoProfile -File .\tests\Probar-Receptor.ps1
powershell.exe -NoProfile -File .\tests\Probar-Configurador.ps1
powershell.exe -NoProfile -File .\tests\Probar-Sesion.ps1
```

`tests/Probar-Inventario.ps1` verifica las funciones de la versión directa, conservada en `Inventario-Graph.ps1`. Ese archivo es la alternativa anterior y sí pide autenticación por computadora; el comando actual generado utiliza el cliente centralizado.

En el entorno restringido de desarrollo no se pudieron verificar consultas CIM reales, la configuración de red de Windows ni una inserción real en Microsoft 365. La puesta en marcha se confirma cuando el central muestre Receptor activo y una captura real termine como RegistradoEnExcel.

Consulte [la conexión con Excel](docs/Configurar-Excel.md) y [la operación del receptor](docs/Receptor-Central.md) para los detalles.
