#requires -Version 5.1
<#
.SYNOPSIS
Recopila el inventario y lo envia al equipo central sin iniciar sesion en Microsoft.
.DESCRIPTION
El receptor central es quien escribe en Excel. Guarda respaldos y una solicitud con ID
estable antes de enviar. -SolicitudJson reenvia esa misma solicitud sin crear otra.
#>
[CmdletBinding()]
param(
    [string]$Area,
    [string]$ServidorUrl,
    [string]$ClaveApi,
    [string]$ConfigPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos\cliente.config.json'),
    [string]$DirectorioSalida = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos'),
    [string]$SolicitudJson,
    [ValidateRange(0,120)][int]$EsperarSegundos = 15,
    [switch]$SoloLocal
)
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

function Assert-ReceiverUrl {
    param([string]$Url)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -notin @('http','https') -or $uri.UserInfo -or $uri.Query -or $uri.Fragment -or $uri.AbsolutePath -ne '/') {
        throw 'ServidorUrl debe ser la direccion base del receptor, sin ruta, usuario ni parametros.'
    }
    if ($uri.Scheme -eq 'http') {
        $ipAddress = $null
        $localHost = $uri.Host -eq 'localhost'
        if ([Net.IPAddress]::TryParse($uri.Host, [ref]$ipAddress) -and $ipAddress.AddressFamily -eq [Net.Sockets.AddressFamily]::InterNetwork) {
            $b = $ipAddress.GetAddressBytes()
            $localHost = ($b[0] -eq 10 -or ($b[0] -eq 172 -and $b[1] -ge 16 -and $b[1] -le 31) -or ($b[0] -eq 192 -and $b[1] -eq 168) -or $b[0] -eq 127)
        }
        if (-not $localHost) { throw 'HTTP solo se admite con una IP privada de la red local. Para Internet configure HTTPS.' }
    }
}

function New-InventoryEnvelope {
    param([object]$Record)
    Assert-InventoryRecord $Record
    return [pscustomobject][ordered]@{ schemaVersion = 1; requestId = [guid]::NewGuid().ToString('D'); record = $Record }
}

function Assert-InventoryEnvelope {
    param([object]$Envelope)
    if ($null -eq $Envelope -or @( $Envelope.PSObject.Properties.Name ).Count -ne 3) { throw 'Solicitud no valida.' }
    $schema = Get-OptionalProperty $Envelope 'schemaVersion'
    $requestId = [string](Get-OptionalProperty $Envelope 'requestId')
    $parsed = [guid]::Empty
    if (($schema -isnot [int] -and $schema -isnot [long]) -or $schema -ne 1 -or -not [guid]::TryParseExact($requestId, 'D', [ref]$parsed)) { throw 'Version o identificador de solicitud no valido.' }
    Assert-InventoryRecord (Get-OptionalProperty $Envelope 'record')
}

function Assert-ReceiverResponse {
    param([object]$Response, [string]$RequestId)
    $status = [string](Get-OptionalProperty $Response 'status')
    if ([string](Get-OptionalProperty $Response 'requestId') -ne $RequestId -or $status -notin @('queued','sent','needs-review','inflight')) {
        throw 'El receptor no devolvio un estado valido para esta solicitud.'
    }
}

function Send-InventoryEnvelope {
    param([object]$Envelope, [string]$Url, [string]$ApiKey, [int]$WaitSeconds = 15)
    Assert-InventoryEnvelope $Envelope
    Assert-ReceiverUrl $Url
    if ([string]::IsNullOrWhiteSpace($ApiKey) -or $ApiKey.Length -lt 32 -or $ApiKey -match '[\r\n]') { throw 'Falta la clave del receptor. Use el comando generado en el equipo central.' }
    $baseUri = $Url.TrimEnd('/')
    $headers = @{ Authorization = 'Bearer ' + $ApiKey; Accept = 'application/json' }
    $json = $Envelope | ConvertTo-Json -Depth 6 -Compress
    # Un solo POST. Si necesita repetirlo, reutilice -SolicitudJson y su requestId.
    $response = Invoke-RestMethod -Method Post -Uri ($baseUri + '/inventario') -Headers $headers -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($json)) -TimeoutSec 20 -MaximumRedirection 0 -ErrorAction Stop
    Assert-ReceiverResponse $response $Envelope.requestId
    $watch = [Diagnostics.Stopwatch]::StartNew()
    while ($response.status -in @('queued','inflight') -and $watch.Elapsed.TotalSeconds -lt $WaitSeconds) {
        Start-Sleep -Milliseconds 1000
        try {
            $next = Invoke-RestMethod -Method Get -Uri ($baseUri + '/estado/' + $Envelope.requestId) -Headers $headers -TimeoutSec 10 -MaximumRedirection 0 -ErrorAction Stop
            Assert-ReceiverResponse $next $Envelope.requestId
            $response = $next
        } catch {
            Write-Verbose 'El receptor ya acepto la solicitud, pero no se pudo consultar su estado actual.'
            break
        }
    }
    return $response
}

function Invoke-InventoryMain {
    [CmdletBinding()]
    param([string]$Area, [string]$ServidorUrl, [string]$ClaveApi, [string]$ConfigPath,
        [string]$DirectorioSalida, [string]$SolicitudJson, [int]$EsperarSegundos = 15, [switch]$SoloLocal)
    $ErrorActionPreference = 'Stop'
    if ($env:OS -ne 'Windows_NT') { throw 'Ejecute el inventario en Windows con PowerShell 5.1 o posterior.' }
    $oldTls = [Net.ServicePointManager]::SecurityProtocol
    $requestPath = $null
    $backup = $null
    try {
        [Net.ServicePointManager]::SecurityProtocol = $oldTls -bor [Net.SecurityProtocolType]::Tls12
        if (-not $SoloLocal -and $ConfigPath -and (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
            $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $ServidorUrl) { $ServidorUrl = [string](Get-OptionalProperty $config 'ServidorUrl') }
            if (-not $ClaveApi) { $ClaveApi = [string](Get-OptionalProperty $config 'ClaveApi') }
        }
        if (-not $ClaveApi) { $ClaveApi = [Environment]::GetEnvironmentVariable('INVENTARIO_API_KEY') }
        if ($SolicitudJson) {
            $envelope = Get-Content -LiteralPath $SolicitudJson -Raw -Encoding UTF8 | ConvertFrom-Json
            Assert-InventoryEnvelope $envelope
            $record = $envelope.record
            $requestPath = (Resolve-Path -LiteralPath $SolicitudJson).ProviderPath
        } else {
            while ([string]::IsNullOrWhiteSpace($Area)) { $Area = Read-Host 'Ingrese el area o departamento (obligatorio)' }
            $record = Get-InventoryRecord -Area $Area
            Assert-InventoryRecord $record
            $envelope = New-InventoryEnvelope $record
        }
        $backup = Save-InventoryBackup $record $DirectorioSalida
        if (-not $requestPath) {
            $requestPath = [IO.Path]::ChangeExtension($backup.Json, 'solicitud.json')
            $envelope | ConvertTo-Json -Depth 6 | Set-Content -LiteralPath $requestPath -Encoding UTF8 -ErrorAction Stop
        }
        Write-Host "Solicitud recuperable: $requestPath"
        if ($SoloLocal) {
            Write-Host 'Inventario guardado. No se enviaron datos.' -ForegroundColor Green
            return [pscustomobject]@{ Estado = 'SoloLocal'; Solicitud = $requestPath; Json = $backup.Json; Csv = $backup.Csv }
        }
        if (-not $ServidorUrl) { throw 'Falta el equipo central. Ejecute el comando que genera Configurar-Servidor.ps1 en ese equipo.' }
        $response = Send-InventoryEnvelope -Envelope $envelope -Url $ServidorUrl -ApiKey $ClaveApi -WaitSeconds $EsperarSegundos
        if ($response.status -eq 'sent') {
            Write-Host 'EXITO: el equipo central confirmo el registro en Excel.' -ForegroundColor Green
            $state = 'RegistradoEnExcel'
        } elseif ($response.status -eq 'needs-review') {
            throw 'El receptor requiere revisar este envio. No repita el registro con otro ID; revise el estado en el equipo central.'
        } else {
            Write-Host 'RECIBIDO: el equipo central guardo la solicitud y la tiene pendiente de registrar en Excel.' -ForegroundColor Yellow
            $state = 'EnCola'
        }
        return [pscustomobject]@{ Estado = $state; SolicitudId = $envelope.requestId; Solicitud = $requestPath; Json = $backup.Json; Csv = $backup.Csv }
    } catch {
        if ($requestPath) { Write-Warning "Conserve la solicitud para consultar o reenviar el mismo ID: $requestPath" }
        throw
    } finally {
        [Net.ServicePointManager]::SecurityProtocol = $oldTls
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $inventoryArguments = @{} + $PSBoundParameters
    if (-not $inventoryArguments.ContainsKey('ConfigPath')) { $inventoryArguments.ConfigPath = $ConfigPath }
    if (-not $inventoryArguments.ContainsKey('DirectorioSalida')) { $inventoryArguments.DirectorioSalida = $DirectorioSalida }
    Invoke-InventoryMain @inventoryArguments
}
