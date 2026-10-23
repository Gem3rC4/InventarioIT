# Conectar el equipo central con Excel

La conexión Microsoft se realiza **solo en el receptor central**. Los clientes ejecutan Inventario-Equipo.ps1 y envían al receptor sin abrir sesión de Microsoft. Siga el README para generar su comando específico.

1. Con la cuenta del central, confirme que puede editar el archivo `.xlsx` de SharePoint.
2. Confirme que los 18 encabezados están dentro de una tabla de Excel real, con estos nombres exactos:

```text
Equipo | Usuario | Marca | Modelo | Serie | Procesador | TipoEquipo | RAM_GB | DiscoTotal_GB | DiscoUsado_GB | DiscoLibre_GB | TipoDisco | IP | Windows | Version | Activado | Fecha | Area
```

3. Copie el vínculo completo del archivo. El configurador lo guarda únicamente en servidor.config.json fuera del repositorio.

## Excel en OneDrive personal

También se admite un Excel de una cuenta personal de Microsoft (outlook.com, hotmail.com). Use cualquiera de estas direcciones como `-ExcelUrl`:

- El enlace para compartir `https://1drv.ms/x/c/...`.
- La dirección del navegador al abrir el archivo en Excel para la web: `https://excel.cloud.microsoft/open/onedrive/?docId=...&driveId=...`. Es la opción más directa, porque apunta al archivo sin resolver el enlace.

El receptor detecta que es OneDrive personal e inicia sesión con el tenant `consumers`: entre con la cuenta personal dueña del archivo. Esto evita las políticas de la cuenta de empresa, pero el receptor debe ejecutarse en un equipo cuyo firewall acepte conexiones de la red local (un cortafuegos de terceros como ESET puede bloquearlas aunque exista la regla de Windows).
4. Si hay más de una tabla, indique su nombre mediante `-Tabla` al configurar. Si hay una sola, se selecciona automáticamente.
5. Inicie el receptor y siga el inicio de sesión por código de dispositivo que muestra Microsoft. Espere a ver Receptor activo.

Configure Serie, IP y Version como Texto para conservar ceros iniciales. Fecha conserva la hora local de la captura, aunque el central procese el registro después. Las cuatro columnas `_GB` reciben números o celdas vacías.

Aunque un vínculo Cualquiera permita editar anónimamente en el navegador, la API de Graph de SharePoint requiere autenticación. Por ello el receptor mantiene la conexión autorizada con Microsoft. `Files.ReadWrite` es un permiso delegado: si el tenant pide aprobación de administrador o bloquea el código de dispositivo, el administrador debe revisar esas políticas. La API de agregar filas a Excel no admite permisos de aplicación.

El receptor no usa el módulo Microsoft.Graph.Authentication (falla en Windows PowerShell 5.1). Inicia sesión con el flujo de código de dispositivo de Microsoft mediante REST, usando el cliente público de Microsoft Graph PowerShell, y guarda el token de renovación en `sesion-microsoft.dat`, cifrado con DPAPI para la cuenta de Windows del central. Renueva el token automáticamente; las políticas, una revocación o una sesión expirada pueden hacer que vuelva a pedir un código. Los clientes no reciben ese token ni una contraseña de Microsoft.

El archivo Inventario-Graph.ps1 conserva la alternativa antigua de conexión directa desde cada computadora. No es el cliente recomendado para el requisito de enviar sin iniciar sesión; el comando generado por Configurar-Servidor.ps1 usa Inventario-Equipo.ps1.

## Fuentes oficiales

- [Instalación de Microsoft Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation).
- [Autenticación y alcance de la caché](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands).
- [Acceder a archivos compartidos](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0).
- [Agregar filas y permisos admitidos](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-add?view=graph-rest-1.0).
