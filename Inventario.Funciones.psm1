#requires -Version 5.1
Set-StrictMode -Version 2.0

function Get-InventoryFields {
    return @('Equipo','Usuario','Marca','Modelo','Serie','Procesador','TipoEquipo','RAM_GB',
        'DiscoTotal_GB','DiscoUsado_GB','DiscoLibre_GB','TipoDisco','IP','Windows','Version','Activado','Fecha','Area')
}

function Get-OptionalProperty {
    param([object]$InputObject, [string]$Name)
    if ($null -eq $InputObject) { return $null }
    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function Get-CimData {
    param([string]$ClassName, [string]$Filter)
    try {
        $parameters = @{ ClassName = $ClassName; ErrorAction = 'Stop' }
        if ($Filter) { $parameters.Filter = $Filter }
        Get-CimInstance @parameters
    } catch { Write-Verbose ("No disponible: {0}." -f $ClassName) }
}

function ConvertTo-Gigabytes {
    param([object]$Bytes)
    if ($null -eq $Bytes) { return $null }
    return [math]::Round(([double]$Bytes / 1GB), 2)
}

function Get-InventoryRecord {
    param([Parameter(Mandatory = $true)][string]$Area)
    if ([string]::IsNullOrWhiteSpace($Area)) { throw 'El area es obligatoria.' }
    $computer = Get-CimData 'Win32_ComputerSystem'
    $os = Get-CimData 'Win32_OperatingSystem'
    if (-not $computer -or -not $os) { throw 'No se pudieron consultar los datos basicos del sistema.' }
    $bios = Get-CimData 'Win32_BIOS'
    $processors = @(Get-CimData 'Win32_Processor')
    $enclosures = @(Get-CimData 'Win32_SystemEnclosure')
    $volumes = @(Get-CimData 'Win32_LogicalDisk' 'DriveType = 3' | Where-Object { $null -ne $_.Size })
    $adapters = @(Get-CimData 'Win32_NetworkAdapterConfiguration' 'IPEnabled = TRUE')
    $hostnameValue = if ($computer.Name) { $computer.Name } else { $env:COMPUTERNAME }
    $userName = if ($computer.UserName) { ($computer.UserName -split '\\')[-1] } else { 'Sin sesion interactiva' }
    $equipmentType = 'Desconocido'
    $chassisTypes = @($enclosures | ForEach-Object { $_.ChassisTypes })
    if (@($chassisTypes | Where-Object { $_ -in @(8,9,10,11,12,14,18,21,30,31,32) }).Count -gt 0) { $equipmentType = 'Laptop' }
    elseif (@($chassisTypes | Where-Object { $_ -in @(3,4,5,6,7,15,16,35,36) }).Count -gt 0) { $equipmentType = 'Desktop' }
    $diskTotal = [double]0
    $diskFree = [double]0
    foreach ($volume in $volumes) {
        $diskTotal += [double]$volume.Size
        if ($null -ne $volume.FreeSpace) { $diskFree += [double]$volume.FreeSpace }
    }
    $diskType = 'Desconocido'
    if (Get-Command Get-PhysicalDisk -ErrorAction SilentlyContinue) {
        try {
            $types = @(Get-PhysicalDisk -ErrorAction Stop | Where-Object { [string]$_.BusType -notin @('USB','SD','MMC') } |
                ForEach-Object { [string]$_.MediaType } | Where-Object { $_ -in @('SSD','HDD') } | Sort-Object -Unique)
            if ($types.Count -gt 0) { $diskType = $types -join '; ' }
        } catch { Write-Verbose 'No se pudo detectar el tipo de disco.' }
    }
    $ips = @($adapters | ForEach-Object { $_.IPAddress } | Where-Object {
        $_ -match '^\d{1,3}(\.\d{1,3}){3}$' -and $_ -notmatch '^(127\.|169\.254\.)'
    } | Sort-Object -Unique)
    $windowsVersion = [string]$os.Version
    try {
        $info = Get-ItemProperty 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
        $displayVersion = Get-OptionalProperty $info 'DisplayVersion'
        $releaseId = Get-OptionalProperty $info 'ReleaseId'
        if ($displayVersion) { $windowsVersion = $displayVersion }
        elseif ($releaseId) { $windowsVersion = $releaseId }
    } catch { Write-Verbose 'Se utilizara la version informada por Windows.' }
    $activation = 'Desconocido'
    $licenses = @(Get-CimData 'SoftwareLicensingProduct' "ApplicationID='55c92734-d682-4d71-983e-d6ec3f16059f' AND PartialProductKey IS NOT NULL")
    if ($licenses.Count -gt 0) {
        $activation = if (@($licenses | Where-Object { $_.LicenseStatus -eq 1 }).Count -gt 0) { 'Activado' } else { 'No activado' }
    }
    $processorNames = @($processors | Where-Object { $_.Name } | ForEach-Object { $_.Name.Trim() } | Sort-Object -Unique)
    [pscustomobject][ordered]@{
        Equipo = $hostnameValue
        Usuario = $userName
        Marca = if ($computer.Manufacturer) { $computer.Manufacturer.Trim() } else { 'Desconocido' }
        Modelo = if ($computer.Model) { $computer.Model.Trim() } else { 'Desconocido' }
        Serie = if ($bios -and $bios.SerialNumber) { $bios.SerialNumber.Trim() } else { 'Desconocido' }
        Procesador = if ($processorNames.Count) { $processorNames -join '; ' } else { 'Desconocido' }
        TipoEquipo = $equipmentType
        RAM_GB = ConvertTo-Gigabytes $computer.TotalPhysicalMemory
        DiscoTotal_GB = if ($volumes.Count) { ConvertTo-Gigabytes $diskTotal } else { $null }
        DiscoUsado_GB = if ($volumes.Count) { ConvertTo-Gigabytes ($diskTotal - $diskFree) } else { $null }
        DiscoLibre_GB = if ($volumes.Count) { ConvertTo-Gigabytes $diskFree } else { $null }
        TipoDisco = $diskType
        IP = if ($ips.Count) { $ips -join '; ' } else { 'Desconocido' }
        Windows = if ($os.Caption) { $os.Caption.Trim() } else { 'Desconocido' }
        Version = if ($windowsVersion) { $windowsVersion } else { 'Desconocido' }
        Activado = $activation
        Fecha = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        Area = $Area.Trim()
    }
}

function Assert-InventoryRecord {
    param([Parameter(Mandatory = $true)][object]$Record)
    $fields = @(Get-InventoryFields)
    $names = @($Record.PSObject.Properties.Name)
    if ($names.Count -ne $fields.Count) { throw 'El registro debe tener exactamente 18 campos.' }
    foreach ($field in $fields) {
        if ($names -cnotcontains $field) { throw "Falta el campo: $field" }
        $value = $Record.PSObject.Properties[$field].Value
        if ($null -ne $value -and $value -isnot [string] -and $value -isnot [ValueType]) { throw "Campo no escalar: $field" }
        if ($field -in @('RAM_GB','DiscoTotal_GB','DiscoUsado_GB','DiscoLibre_GB')) {
            if ($null -ne $value -and ($value -is [string] -or $value -is [bool] -or [double]$value -lt 0 -or
                [double]::IsNaN([double]$value) -or [double]::IsInfinity([double]$value))) { throw "Numero no valido: $field" }
        } elseif ($null -eq $value -or $value -isnot [string] -or [string]::IsNullOrWhiteSpace($value)) {
            throw "Texto obligatorio no valido: $field"
        }
    }
}

function ConvertTo-ExcelSafeValue {
    param([AllowNull()][object]$Value)
    # Excel no debe interpretar nombres, series o areas como formulas.
    if ($Value -is [string] -and $Value -match '^[\s]*[=+@-]') { return "'" + $Value }
    return $Value
}

function Save-InventoryBackup {
    param([object]$Record, [string]$Directory)
    $null = New-Item -ItemType Directory -Path $Directory -Force -ErrorAction Stop
    $safeName = $Record.Equipo -replace '[^A-Za-z0-9._-]', '_'
    $baseName = '{0}_{1}_{2}' -f $safeName, (Get-Date -Format 'yyyyMMdd_HHmmss_fff'), ([guid]::NewGuid().ToString('N').Substring(0,8))
    $jsonPath = Join-Path $Directory ($baseName + '.json')
    $csvPath = Join-Path $Directory ($baseName + '.csv')
    $Record | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $jsonPath -Encoding UTF8 -ErrorAction Stop
    Write-Host "Respaldo JSON: $jsonPath"
    $safeRecord = [ordered]@{}
    foreach ($field in Get-InventoryFields) { $safeRecord[$field] = ConvertTo-ExcelSafeValue $Record.PSObject.Properties[$field].Value }
    [pscustomobject]$safeRecord | Export-Csv -LiteralPath $csvPath -NoTypeInformation -Encoding UTF8 -ErrorAction Stop
    Write-Host "Respaldo CSV:  $csvPath"
    return [pscustomobject]@{ Json = $jsonPath; Csv = $csvPath }
}

function ConvertTo-SharingToken {
    param([string]$Url)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { throw 'ExcelUrl debe ser un enlace HTTPS completo.' }
    if ($uri.Host -notmatch '(^|\.)sharepoint\.com$' -and $uri.Host -notin @('1drv.ms','onedrive.live.com')) {
        throw 'ExcelUrl debe pertenecer a SharePoint o OneDrive de Microsoft.'
    }
    return 'u!' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Url)).TrimEnd('=').Replace('/','_').Replace('+','-')
}

# ---------------------------------------------------------------------------
# Sesion de Microsoft del receptor central, sin el modulo Microsoft.Graph.
# Ese modulo falla en Windows PowerShell 5.1 ("Error al escribir en un agente de
# escucha"), por eso se usa el flujo de codigo de dispositivo con REST directo.
# ---------------------------------------------------------------------------
$script:GraphSession = $null
$script:GraphDefaultClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'  # Microsoft Graph PowerShell (cliente publico)
$script:GraphScope = 'https://graph.microsoft.com/Files.ReadWrite offline_access openid profile'

function Get-HttpErrorInfo {
    param([Parameter(Mandatory = $true)][Management.Automation.ErrorRecord]$ErrorRecord)
    $statusCode = $null
    $response = Get-OptionalProperty $ErrorRecord.Exception 'Response'
    if ($null -ne $response) { try { $statusCode = [int]$response.StatusCode } catch { $statusCode = $null } }
    $text = $null
    if ($null -ne $ErrorRecord.ErrorDetails -and $ErrorRecord.ErrorDetails.Message) { $text = $ErrorRecord.ErrorDetails.Message }
    elseif ($null -ne $response) {
        try { $reader = New-Object IO.StreamReader($response.GetResponseStream()); $text = $reader.ReadToEnd(); $reader.Dispose() } catch { $text = $null }
    }
    $code = $null; $message = $ErrorRecord.Exception.Message
    if ($text) {
        $message = $text
        try {
            $json = $text | ConvertFrom-Json -ErrorAction Stop
            $errorValue = Get-OptionalProperty $json 'error'
            if ($errorValue -is [string]) {
                # Formato OAuth: error / error_description
                $code = $errorValue
                $description = [string](Get-OptionalProperty $json 'error_description')
                if ($description) { $message = ($description -split '\r?\n')[0] }
            } elseif ($null -ne $errorValue) {
                # Formato Graph: error.code / error.message
                $code = [string](Get-OptionalProperty $errorValue 'code')
                $message = [string](Get-OptionalProperty $errorValue 'message')
            }
        } catch { $message = $text }
    }
    return [pscustomobject]@{ StatusCode = $statusCode; Code = $code; Message = $message }
}

function Invoke-OAuthRequest {
    param([string]$Tenant, [string]$Endpoint, [hashtable]$Form)
    $uri = 'https://login.microsoftonline.com/{0}/oauth2/v2.0/{1}' -f [uri]::EscapeDataString($Tenant), $Endpoint
    return Invoke-RestMethod -Method Post -Uri $uri -Body $Form -ContentType 'application/x-www-form-urlencoded' -TimeoutSec 60 -UseBasicParsing -ErrorAction Stop
}

function Save-GraphRefreshToken {
    param([string]$CachePath, [string]$Tenant, [string]$ClientId, [string]$RefreshToken)
    if (-not $CachePath -or -not $RefreshToken) { return }
    try {
        # DPAPI: solo esta cuenta de Windows en este equipo puede descifrarlo.
        $protected = ConvertTo-SecureString -String $RefreshToken -AsPlainText -Force | ConvertFrom-SecureString
        $content = [ordered]@{ Tenant = $Tenant; ClientId = $ClientId; Protegido = $protected } | ConvertTo-Json
        [IO.File]::WriteAllText($CachePath, $content, (New-Object Text.UTF8Encoding($false)))
    } catch { Write-Warning 'No se pudo guardar la sesion de Microsoft; al reiniciar el receptor se pedira un codigo nuevo.' }
}

function Read-GraphRefreshToken {
    param([string]$CachePath, [string]$Tenant, [string]$ClientId)
    if (-not $CachePath -or -not (Test-Path -LiteralPath $CachePath -PathType Leaf)) { return $null }
    $cached = [IO.File]::ReadAllText($CachePath, [Text.Encoding]::UTF8) | ConvertFrom-Json
    if ([string](Get-OptionalProperty $cached 'Tenant') -ne $Tenant -or [string](Get-OptionalProperty $cached 'ClientId') -ne $ClientId) { return $null }
    $secure = ConvertTo-SecureString -String ([string](Get-OptionalProperty $cached 'Protegido'))
    $pointer = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try { return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer) }
}

function Set-GraphSessionToken {
    param([object]$Token, [string]$Tenant, [string]$ClientId, [string]$CachePath)
    $accessToken = [string](Get-OptionalProperty $Token 'access_token')
    if (-not $accessToken) { throw 'Microsoft no devolvio un token de acceso.' }
    $refreshToken = [string](Get-OptionalProperty $Token 'refresh_token')
    if (-not $refreshToken -and $null -ne $script:GraphSession) { $refreshToken = $script:GraphSession.RefreshToken }
    $expiresIn = 3600
    $expiresValue = Get-OptionalProperty $Token 'expires_in'
    if ($null -ne $expiresValue) { $expiresIn = [int]$expiresValue }
    $script:GraphSession = [pscustomobject]@{
        AccessToken = $accessToken; RefreshToken = $refreshToken
        ExpiresAtUtc = [DateTime]::UtcNow.AddSeconds([math]::Max(60, $expiresIn - 300))
        Tenant = $Tenant; ClientId = $ClientId; CachePath = $CachePath
    }
    Save-GraphRefreshToken -CachePath $CachePath -Tenant $Tenant -ClientId $ClientId -RefreshToken $refreshToken
}

function Connect-InventarioGraph {
    <#
    Inicia sesion delegada (Files.ReadWrite) con codigo de dispositivo. Reutiliza la sesion
    guardada si sigue valida, para no pedir el codigo en cada reinicio del receptor.
    #>
    param([string]$TenantId, [string]$ClientId, [string]$CachePath, [switch]$NuevaSesion)
    $tenant = if ($TenantId) { $TenantId } else { 'organizations' }
    $client = if ($ClientId) { $ClientId } else { $script:GraphDefaultClientId }
    $script:GraphSession = $null
    if (-not $NuevaSesion) {
        try {
            $refreshToken = Read-GraphRefreshToken -CachePath $CachePath -Tenant $tenant -ClientId $client
            if ($refreshToken) {
                $token = Invoke-OAuthRequest $tenant 'token' @{ client_id = $client; grant_type = 'refresh_token'; refresh_token = $refreshToken; scope = $script:GraphScope }
                Set-GraphSessionToken -Token $token -Tenant $tenant -ClientId $client -CachePath $CachePath
                Write-Host 'Sesion de Microsoft recuperada; no hace falta un codigo nuevo.' -ForegroundColor Green
                return
            }
        } catch { Write-Host 'La sesion guardada de Microsoft ya no es valida. Se pedira iniciar sesion.' -ForegroundColor Yellow }
    }
    try { $device = Invoke-OAuthRequest $tenant 'devicecode' @{ client_id = $client; scope = $script:GraphScope } }
    catch { $info = Get-HttpErrorInfo $_; throw ('No se pudo solicitar el codigo de inicio de sesion a Microsoft: ' + $info.Message) }
    $userCode = [string]$device.user_code
    $verificationUri = [string]$device.verification_uri
    Write-Host ''
    Write-Host '================ INICIO DE SESION DE MICROSOFT (solo en el central) ================' -ForegroundColor Cyan
    Write-Host ('  1. Abra en el navegador:  ' + $verificationUri) -ForegroundColor Cyan
    Write-Host ('  2. Escriba este codigo:   ' + $userCode) -ForegroundColor Yellow
    Write-Host '  3. Inicie sesion con la cuenta que puede editar el Excel y acepte los permisos.' -ForegroundColor Cyan
    Write-Host '====================================================================================' -ForegroundColor Cyan
    try { Set-Clipboard -Value $userCode; Write-Host '(El codigo ya esta copiado en el portapapeles.)' } catch { Write-Verbose 'Sin portapapeles.' }
    try { Start-Process $verificationUri } catch { Write-Verbose 'No se pudo abrir el navegador.' }
    $interval = 5
    $intervalValue = Get-OptionalProperty $device 'interval'
    if ($null -ne $intervalValue) { $interval = [math]::Max(1, [int]$intervalValue) }
    $expiresIn = 900
    $expiresValue = Get-OptionalProperty $device 'expires_in'
    if ($null -ne $expiresValue) { $expiresIn = [int]$expiresValue }
    $deadline = [DateTime]::UtcNow.AddSeconds($expiresIn)
    Write-Host 'Esperando que complete el inicio de sesion...'
    while ([DateTime]::UtcNow -lt $deadline) {
        Start-Sleep -Seconds $interval
        try {
            $token = Invoke-OAuthRequest $tenant 'token' @{ client_id = $client; grant_type = 'urn:ietf:params:oauth:grant-type:device_code'; device_code = [string]$device.device_code }
        } catch {
            $info = Get-HttpErrorInfo $_
            if ($info.Code -eq 'authorization_pending') { continue }
            if ($info.Code -eq 'slow_down') { $interval += 5; continue }
            if ($info.Code -eq 'authorization_declined') { throw 'El inicio de sesion fue cancelado en el navegador.' }
            if ($info.Code -eq 'expired_token') { break }
            throw ('Microsoft rechazo el inicio de sesion: ' + $info.Message)
        }
        Set-GraphSessionToken -Token $token -Tenant $tenant -ClientId $client -CachePath $CachePath
        Write-Host 'Inicio de sesion de Microsoft completado.' -ForegroundColor Green
        return
    }
    throw 'El codigo de inicio de sesion expiro. Vuelva a iniciar el receptor y complete el inicio de sesion en 15 minutos.'
}

function Get-InventarioGraphAccessToken {
    param([switch]$Force)
    if ($null -eq $script:GraphSession) { throw 'El receptor no tiene sesion de Microsoft. Reinicie el receptor.' }
    if ($Force -or [DateTime]::UtcNow -ge $script:GraphSession.ExpiresAtUtc) {
        if (-not $script:GraphSession.RefreshToken) { throw 'La sesion de Microsoft expiro. Reinicie el receptor.' }
        $session = $script:GraphSession
        try {
            $token = Invoke-OAuthRequest $session.Tenant 'token' @{ client_id = $session.ClientId; grant_type = 'refresh_token'; refresh_token = $session.RefreshToken; scope = $script:GraphScope }
        } catch { $info = Get-HttpErrorInfo $_; throw ('No se pudo renovar la sesion de Microsoft; reinicie el receptor. Detalle: ' + $info.Message) }
        Set-GraphSessionToken -Token $token -Tenant $session.Tenant -ClientId $session.ClientId -CachePath $session.CachePath
    }
    return $script:GraphSession.AccessToken
}

function Invoke-InventarioGraphRequest {
    param([string]$Method = 'GET', [Parameter(Mandatory = $true)][string]$Uri, [AllowNull()][string]$Body)
    # No reenviar el token a otro host.
    if (-not $Uri.StartsWith('https://graph.microsoft.com/v1.0/', [StringComparison]::OrdinalIgnoreCase)) { throw 'URL de Graph no permitida.' }
    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $headers = @{ Authorization = 'Bearer ' + (Get-InventarioGraphAccessToken -Force:($attempt -gt 1)); Accept = 'application/json' }
        $parameters = @{ Method = $Method; Uri = $Uri; Headers = $headers; TimeoutSec = 120; UseBasicParsing = $true; ErrorAction = 'Stop' }
        if ($null -ne $Body -and $Body -ne '') { $parameters.Body = [Text.Encoding]::UTF8.GetBytes($Body); $parameters.ContentType = 'application/json; charset=utf-8' }
        try { return Invoke-RestMethod @parameters }
        catch {
            $info = Get-HttpErrorInfo $_
            # 401: token rechazado antes de ejecutar; se renueva y se reintenta una sola vez.
            if ($info.StatusCode -eq 401 -and $attempt -eq 1) { continue }
            throw ('Microsoft Graph respondio {0} {1}: {2}' -f $info.StatusCode, $info.Code, $info.Message)
        }
    }
}

function Get-GraphCollection {
    param([string]$Uri)
    $items = New-Object 'System.Collections.Generic.List[object]'
    do {
        # No reenviar credenciales a otro host al seguir paginacion.
        if (-not $Uri.StartsWith('https://graph.microsoft.com/v1.0/', [StringComparison]::OrdinalIgnoreCase)) { throw 'URL de Graph no permitida.' }
        $page = Invoke-InventarioGraphRequest -Method GET -Uri $Uri
        if ($null -eq $page.PSObject.Properties['value']) { throw 'Respuesta de coleccion de Graph no valida.' }
        foreach ($item in @($page.value)) { if ($null -ne $item) { $items.Add($item) } }
        $Uri = [string](Get-OptionalProperty $page '@odata.nextLink')
    } while ($Uri)
    return $items.ToArray()
}

function Get-ExcelDestination {
    param([string]$Url, [string]$TableName)
    $token = ConvertTo-SharingToken $Url
    try { $item = Invoke-InventarioGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/shares/$token/driveItem" }
    catch { throw ('No se pudo abrir el Excel con la cuenta de Microsoft del central. Inicie sesion con una cuenta que pueda editarlo. ' + $_.Exception.Message) }
    $remote = Get-OptionalProperty $item 'remoteItem'
    if ($remote) { $item = $remote }
    $itemId = [string](Get-OptionalProperty $item 'id')
    $driveId = [string](Get-OptionalProperty (Get-OptionalProperty $item 'parentReference') 'driveId')
    if (-not $itemId -or -not $driveId) { throw 'Graph no devolvio los identificadores del archivo. Abra el enlace con la cuenta que tiene permiso de edicion.' }
    $base = 'https://graph.microsoft.com/v1.0/drives/{0}/items/{1}/workbook' -f [uri]::EscapeDataString($driveId), [uri]::EscapeDataString($itemId)
    $tables = @(Get-GraphCollection "$base/tables")
    if ($TableName) {
        $matches = @($tables | Where-Object { $_.name -eq $TableName -or $_.id -eq $TableName })
    } else { $matches = $tables }
    if ($matches.Count -ne 1) {
        $available = ($tables | ForEach-Object { $_.name }) -join ', '
        throw "No se pudo seleccionar una sola tabla. Use -Tabla con el nombre exacto. Disponibles: $available"
    }
    $table = $matches[0]
    $tableUri = "$base/tables/$([uri]::EscapeDataString([string]$table.id))"
    $columns = @(Get-GraphCollection "$tableUri/columns" | Sort-Object { [int]$_.index })
    $fields = @(Get-InventoryFields)
    if ($columns.Count -ne $fields.Count) { throw 'La tabla debe contener exactamente las 18 columnas del inventario. No se envio ninguna fila.' }
    $seen = @()
    for ($i = 0; $i -lt $columns.Count; $i++) {
        $name = [string]$columns[$i].name
        if ($fields -cnotcontains $name -or $seen -contains $name -or [int]$columns[$i].index -ne $i) {
            throw "Columna inesperada o repetida: $name. Revise los encabezados exactos; no se envio ninguna fila."
        }
        $seen += $name
    }
    return [pscustomobject]@{ TableUri = $tableUri; TableName = $table.name; Columns = $columns }
}

function New-InventoryRowBody {
    param([object]$Record, [object[]]$Columns)
    Assert-InventoryRecord $Record
    $row = @($Columns | ForEach-Object { ConvertTo-ExcelSafeValue $Record.PSObject.Properties[[string]$_.name].Value })
    # Matriz 1 x 18 explicita: PowerShell tiende a aplanar arrays.
    $matrix = [object[]]::new(1)
    $matrix[0] = $row
    return (@{ values = $matrix } | ConvertTo-Json -Depth 6 -Compress)
}

function Initialize-GraphModule {
    param([switch]$Install)
    if (-not (Get-Module -ListAvailable Microsoft.Graph.Authentication)) {
        if (-not $Install) { throw 'Falta Microsoft.Graph.Authentication. Ejecute de nuevo con -InstalarDependencias o consulte README.md.' }
        Write-Host 'Instalando Microsoft.Graph.Authentication para el usuario actual desde PSGallery...'
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Repository PSGallery -MinimumVersion 2.0.0 -Force -AllowClobber -ErrorAction Stop
    }
    Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -ErrorAction Stop
}

function Get-UrlAclState {
    <#
    Interpreta la salida de "netsh http show urlacl url=<prefijo>" sin depender del idioma.
    Solo usa el prefijo y la linea SDDL, que no se traducen.
    Devuelve: Missing (sin reserva), Owned (permite escuchar al SID) o Foreign (otra identidad).
    #>
    param([AllowNull()][string]$NetshOutput, [Parameter(Mandatory = $true)][string]$UrlPrefix, [Parameter(Mandatory = $true)][string]$Sid)
    if ([string]::IsNullOrWhiteSpace($NetshOutput)) { return 'Missing' }
    $prefix = $UrlPrefix.Trim().TrimEnd('/') + '/'
    $lines = @($NetshOutput -split '\r?\n')
    $found = $false
    $sddl = New-Object 'System.Collections.Generic.List[string]'
    foreach ($line in $lines) {
        $trimmed = $line.Trim()
        # Linea "Reserved URL : http://..." o "Direccion URL reservada : http://...".
        $lineUrl = $null
        if ($trimmed -match '(https?://\S+)') { $lineUrl = $Matches[1].TrimEnd('/') + '/' }
        if (-not $found) {
            if ($lineUrl -and $lineUrl -ieq $prefix) { $found = $true }
            continue
        }
        if ($lineUrl) { break }
        if ($trimmed -match '^SDDL\s*:\s*(\S+)') { $sddl.Add($Matches[1]) }
    }
    if (-not $found) { return 'Missing' }
    # Windows abrevia algunas cuentas en el SDDL (LA = Administrador integrado, BU = Usuarios...).
    # RawSecurityDescriptor convierte esas abreviaturas al SID real antes de comparar.
    $allowed = $false
    foreach ($descriptor in $sddl) {
        try { $acl = ([Security.AccessControl.RawSecurityDescriptor]::new($descriptor)).DiscretionaryAcl }
        catch { continue }
        if ($null -eq $acl) { continue }
        foreach ($ace in $acl) {
            if ($ace -isnot [Security.AccessControl.CommonAce] -or $ace.SecurityIdentifier.Value -ine $Sid) { continue }
            # 0x20000000 = GX (escuchar), 0x10000000 = GA (todo).
            if (($ace.AccessMask -band 0x30000000) -eq 0) { continue }
            if ($ace.AceQualifier -eq [Security.AccessControl.AceQualifier]::AccessDenied) { return 'Foreign' }
            if ($ace.AceQualifier -eq [Security.AccessControl.AceQualifier]::AccessAllowed) { $allowed = $true }
        }
    }
    if ($allowed) { return 'Owned' }
    return 'Foreign'
}

Export-ModuleMember -Function *
