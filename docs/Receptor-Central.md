# Receptor central de InventarioIT

## Inicio y archivos

`Configurar-Servidor.ps1` genera servidor.config.json, cliente.config.json, Comando-Equipos.txt e Iniciar-Servidor.cmd en `%LOCALAPPDATA%\InventarioEquipos`. La cola se guarda en su subcarpeta `cola`, con un archivo GUID.json por solicitud. No publique estos archivos ni los copie al repositorio.

El configurador no habilita recepción hasta iniciar el servidor. Al usar `-PrepararRed`, necesita privilegios de administrador para reservar exactamente la IP y el puerto en HTTP.sys y habilitarlos únicamente para la subred local en perfiles Dominio/Privado. No cambia perfiles ni reglas ajenas. Si hay una reserva para otra identidad o una regla con el mismo nombre que no pertenece a InventarioIT, se detiene para que el administrador la revise.

El receptor debe ejecutarse como la misma cuenta Windows de la configuración. No necesita mantenerse elevado después de reservar la URL, siempre que esa cuenta tenga acceso a su configuración y cola.

El central inicia sesión con Microsoft, comprueba el Excel y su tabla, y después abre la recepción. La conexión es delegada `Files.ReadWrite`, con caché CurrentUser local del SDK; puede requerir reautenticación por reinicio, revocación o políticas corporativas. No se distribuye ese token a los clientes.

## Protocolo

- POST `/inventario`, con `Authorization: Bearer CLAVE` y JSON UTF-8: `{schemaVersion:1, requestId:"GUID-con-guiones", record:{18 campos}}`.
- GET `/estado/GUID-con-guiones`, con la misma clave: devuelve estado y fechas, sin el registro de hardware.
- La recepción válida se persiste antes de responder `202`, estado `queued`.
- Un ID existente con igual contenido devuelve el recibo existente; con contenido distinto devuelve `409`.
- Los cuerpos se limitan a 64 KiB y la cola a 1000 pendientes. La API no admite acceso anónimo.

La clave de recepción es una credencial limitada de esta aplicación. El usuario del cliente no inicia sesión en Microsoft, aunque el comando lleva esta clave para autorizar la recepción. Con HTTP local, la clave y los datos no van cifrados; use la modalidad generada solo en una LAN confiable. Para otras redes configure HTTPS con certificado válido en HTTP.sys y no abra el receptor directamente a Internet.

## Escritura y estados

Un solo proceso utiliza cada directorio de cola, protegido por un bloqueo de archivo. El trabajador procesa secuencialmente:

1. `queued`: guardado en el central; aún no insertado.
2. Valida nuevamente el archivo y las 18 columnas.
3. Persiste `inflight` antes de hacer una sola solicitud de inserción a Excel, con reintentos del SDK deshabilitados.
4. `sent`: solo después de recibir el índice confirmado de la fila y guardar el recibo.
5. `needs-review`: error de validación, error de inserción, confirmación incierta o reinicio con un envío en curso.

Los clientes ven `queued` mientras una solicitud está en curso. El central muestra `[sent]` o `[needs-review]` con el identificador, sin imprimir la clave ni los datos completos.

No se reintentan automáticamente los envíos inciertos. Revise `Equipo` y `Fecha` en el Excel y el `lastError` del recibo. Conserve la cola para auditoría y deduplicación. El Excel de 18 columnas no contiene una clave única de operación, por lo que una respuesta perdida no permite demostrar por sí sola si se insertó o no.

## Incidencias comunes

- Cliente sin conexión: compruebe que el central está encendido, la ventana del receptor está abierta y aparece Receptor activo. Revise IP, puerto, perfil de red y firewall. Use una IP fija o reserva DHCP.
- `401`: revise que el comando o cliente.config.json usen la clave generada en ese central.
- `409`: se reutilizó un ID con contenido diferente; no cambie los datos de una solicitud guardada.
- `503`: la cola de pendientes está llena; reintente más tarde con el mismo archivo de solicitud.
- Error de Microsoft o `needs-review`: revise la consola, los permisos del archivo y los recibos. La sesión Microsoft solo se renueva en el central.
- No detenga el proceso cerrando Windows mientras está escribiendo. Ctrl+C libera el listener y el bloqueo; los recibos que quedaron inflight se revisarán al próximo inicio.

Un Graph lento puede bloquear el trabajador hasta su límite de tiempo y retrasar respuestas posteriores, porque esta implementación pequeña usa un único hilo. El cliente conserva la solicitud si su espera termina antes. No está diseñado como receptor masivo de Internet ni como servicio Windows con alta disponibilidad.

## Fuentes oficiales

- [Microsoft Graph para agregar filas](https://learn.microsoft.com/en-us/graph/api/tablerowcollection-add?view=graph-rest-1.0): autenticación delegada y confirmación de inserción.
- [Autenticación de Graph PowerShell](https://learn.microsoft.com/en-us/powershell/microsoftgraph/authentication-commands): caché y alcance CurrentUser.
- [Renovación de tokens](https://learn.microsoft.com/en-us/entra/identity-platform/refresh-tokens): revocación y necesidad de volver a autenticarse.
- [HttpListener](https://learn.microsoft.com/en-us/dotnet/api/system.net.httplistener?view=netframework-4.8.1): prefijos y escucha HTTP de Windows.
- [Configuración HTTP.sys](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/netsh-http): reservas de URL y certificados HTTPS.
