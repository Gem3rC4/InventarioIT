# Conectar el equipo central con Excel

La conexión Microsoft se realiza **solo en el receptor central**. Los clientes ejecutan Inventario-Equipo.ps1 y envían al receptor sin abrir sesión de Microsoft. Siga el README para generar su comando específico.

1. Con la cuenta del central, confirme que puede editar el archivo `.xlsx` de SharePoint.
2. Confirme que los 18 encabezados están dentro de una tabla de Excel real, con estos nombres exactos:

```text
Equipo | Usuario | Marca | Modelo | Serie | Procesador | TipoEquipo | RAM_GB | DiscoTotal_GB | DiscoUsado_GB | DiscoLibre_GB | TipoDisco | IP | Windows | Version | Activado | Fecha | Area
```

3. Copie el vínculo completo del archivo. El configurador lo guarda únicamente en servidor.config.json fuera del repositorio.
4. Si hay más de una tabla, indique su nombre mediante `-Tabla` al configurar. Si hay una sola, se selecciona automáticamente.
5. Inicie el receptor y siga el inicio de sesión por código de dispositivo que muestra Microsoft. Espere a ver Receptor activo.

Configure Serie, IP y Version como Texto para conservar ceros iniciales. Fecha conserva la hora local de la captura, aunque el central procese el registro después. Las cuatro columnas `_GB` reciben números o celdas vacías.

Aunque un vínculo Cualquiera permita editar anónimamente en el navegador, la API de Graph de SharePoint requiere autenticación. Por ello el receptor mantiene la conexión autorizada con Microsoft. `Files.ReadWrite` es un permiso delegado: si el tenant pide aprobación de administrador o bloquea el código de dispositivo, el administrador debe revisar esas políticas. La API de agregar filas a Excel no admite permisos de aplicación.

Microsoft.Graph.Authentication se instala solo en el central. Windows PowerShell 5.1 requiere .NET Framework 4.7.2 o posterior y PowerShellGet actualizado. El SDK conserva su caché en el contexto CurrentUser y puede renovar tokens; las políticas, revocación o reinicios pueden hacer que vuelva a solicitar acceso en ese equipo. Los clientes no reciben ese token ni una contraseña de Microsoft.

El archivo Inventario-Graph.ps1 conserva la alternativa antigua de conexión directa desde cada computadora. No es el cliente recomendado para el requisito de enviar sin iniciar sesión; el comando generado por Configurar-Servidor.ps1 usa Inventario-Equipo.ps1.

## Fuentes oficiales

- [Instalación de Microsoft Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/installation).
- [Autenticación y alcance de la caché](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands).
- [Acceder a archivos compartidos](https://learn.microsoft.com/en-us/graph/api/shares-get?view=graph-rest-1.0).
- [Agregar filas y permisos admitidos](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-add?view=graph-rest-1.0).
