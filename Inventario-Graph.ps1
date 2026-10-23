#requires -Version 5.1
<#
.SYNOPSIS
Recopila los 18 campos de inventario y agrega una fila a Excel Online mediante Microsoft Graph.
.DESCRIPTION
Conserva JSON y CSV locales antes de conectarse. Requiere una cuenta con permiso de
edicion del archivo. No requiere Excel instalado ni Power Automate.
Use -SoloLocal para recopilar sin conectarse o -ComprobarDestino para validar sin insertar.
#>
[CmdletBinding()]
param(
    [string]$Area,
    [string]$ExcelUrl,
    [string]$Tabla,
    [string]$TenantId,
    [string]$ClientId,
    [string]$ConfigPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos\config.json'),
    [string]$DirectorioSalida = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos'),
    [string]$RegistroJson,
    [switch]$SoloLocal,
    [switch]$ComprobarDestino,
    [switch]$InstalarDependencias
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

function ConvertTo-SharingToken {
    param([string]$Url)
    $uri = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$uri) -or $uri.Scheme -ne 'https') { throw 'ExcelUrl debe ser un enlace HTTPS completo.' }
    if ($uri.Host -notmatch '(^|\.)sharepoint\.com$' -and $uri.Host -notin @('1drv.ms','onedrive.live.com')) {
        throw 'ExcelUrl debe pertenecer a SharePoint o OneDrive de Microsoft.'
    }
    return 'u!' + [Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($Url)).TrimEnd('=').Replace('/','_').Replace('+','-')
}

function Get-GraphCollection {
    param([string]$Uri)
    $items = New-Object 'System.Collections.Generic.List[object]'
    do {
        # No reenviar credenciales a otro host al seguir paginacion.
        if (-not $Uri.StartsWith('https://graph.microsoft.com/v1.0/', [StringComparison]::OrdinalIgnoreCase)) { throw 'URL de Graph no permitida.' }
        $page = Invoke-MgGraphRequest -Method GET -Uri $Uri -OutputType PSObject -ErrorAction Stop
        if ($null -eq $page.PSObject.Properties['value']) { throw 'Respuesta de coleccion de Graph no valida.' }
        foreach ($item in @($page.value)) { if ($null -ne $item) { $items.Add($item) } }
        $Uri = [string](Get-OptionalProperty $page '@odata.nextLink')
    } while ($Uri)
    return $items.ToArray()
}

function Get-ExcelDestination {
    param([string]$Url, [string]$TableName)
    $token = ConvertTo-SharingToken $Url
    $item = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/shares/$token/driveItem" -OutputType PSObject -ErrorAction Stop
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

function Invoke-InventoryMain {
    [CmdletBinding()]
    param([string]$Area, [string]$ExcelUrl, [string]$Tabla, [string]$TenantId, [string]$ClientId,
        [string]$ConfigPath, [string]$DirectorioSalida, [string]$RegistroJson,
        [switch]$SoloLocal, [switch]$ComprobarDestino, [switch]$InstalarDependencias)
    $ErrorActionPreference = 'Stop'
    if ($env:OS -ne 'Windows_NT') { throw 'Ejecute este script en Windows con PowerShell 5.1 o posterior.' }
    if ($SoloLocal -and $ComprobarDestino) { throw 'Use SoloLocal o ComprobarDestino, no ambos.' }
    if ($ComprobarDestino -and $RegistroJson) { throw 'ComprobarDestino no acepta RegistroJson.' }
    $previousTls = [Net.ServicePointManager]::SecurityProtocol
    $previousRequestContext = $null
    $connected = $false
    $backup = $null
    $sendAttempted = $false
    try {
        [Net.ServicePointManager]::SecurityProtocol = $previousTls -bor [Net.SecurityProtocolType]::Tls12
        if ($ConfigPath -and (Test-Path -LiteralPath $ConfigPath -PathType Leaf)) {
            $config = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
            if (-not $ExcelUrl) { $ExcelUrl = [string](Get-OptionalProperty $config 'ExcelUrl') }
            if (-not $Tabla) { $Tabla = [string](Get-OptionalProperty $config 'Tabla') }
            if (-not $TenantId) { $TenantId = [string](Get-OptionalProperty $config 'TenantId') }
            if (-not $ClientId) { $ClientId = [string](Get-OptionalProperty $config 'ClientId') }
        }
        if (-not $ComprobarDestino) {
            if ($RegistroJson) {
                $record = Get-Content -LiteralPath $RegistroJson -Raw -Encoding UTF8 | ConvertFrom-Json
            } else {
                while ([string]::IsNullOrWhiteSpace($Area)) { $Area = Read-Host 'Ingrese el area o departamento (obligatorio)' }
                $record = Get-InventoryRecord -Area $Area
            }
            Assert-InventoryRecord $record
            $backup = Save-InventoryBackup $record $DirectorioSalida
            if ($SoloLocal) {
                Write-Host 'Inventario guardado localmente. No se envio a Excel.' -ForegroundColor Green
                return [pscustomobject]@{ Estado = 'SoloLocal'; Json = $backup.Json; Csv = $backup.Csv }
            }
        }
        if (-not $ExcelUrl) { $ExcelUrl = Read-Host 'Pegue el enlace completo del archivo Excel de SharePoint o OneDrive' }
        $null = ConvertTo-SharingToken $ExcelUrl
        Initialize-GraphModule -Install:$InstalarDependencias
        $existingContext = Get-MgContext
        if ($existingContext) { throw 'Ya existe una sesion Graph en esta consola. Abra una nueva ventana de PowerShell para elegir la cuenta del inventario.' }
        $connect = @{ Scopes = @('Files.ReadWrite'); ContextScope = 'Process'; UseDeviceAuthentication = $true; NoWelcome = $true; ErrorAction = 'Stop' }
        if ($TenantId) { $connect.TenantId = $TenantId }
        if ($ClientId) { $connect.ClientId = $ClientId }
        Write-Host 'Inicie sesion con una cuenta Microsoft 365 que tenga permiso de edicion sobre el Excel.' -ForegroundColor Cyan
        $null = Connect-MgGraph @connect
        $connected = $true
        $requestContext = Get-MgRequestContext
        $previousRequestContext = @{
            MaxRetry = [int]$requestContext.MaxRetry
            ClientTimeoutSeconds = [int]$requestContext.ClientTimeout.TotalSeconds
        }
        # Un POST repetido por timeout podria duplicar una fila.
        $null = Set-MgRequestContext -MaxRetry 0 -ClientTimeout 120
        $destination = Get-ExcelDestination -Url $ExcelUrl -TableName $Tabla
        if ($ComprobarDestino) {
            Write-Host "Destino verificado. Tabla: $($destination.TableName). Las 18 columnas coinciden. No se inserto ninguna fila." -ForegroundColor Green
            return [pscustomobject]@{ Estado = 'DestinoVerificado'; Tabla = $destination.TableName }
        }
        $body = New-InventoryRowBody $record $destination.Columns
        $sendAttempted = $true
        $response = Invoke-MgGraphRequest -Method POST -Uri ($destination.TableUri + '/rows/add') -Body $body -ContentType 'application/json; charset=utf-8' -OutputType PSObject -ErrorAction Stop
        if ($null -eq (Get-OptionalProperty $response 'index')) { throw 'No se recibio confirmacion de la fila. Verifique Excel antes de repetir el envio.' }
        Write-Host "EXITO: equipo registrado en Excel. Tabla: $($destination.TableName)." -ForegroundColor Green
        return [pscustomobject]@{ Estado = 'Enviado'; Tabla = $destination.TableName; IndiceFila = $response.index; Json = $backup.Json; Csv = $backup.Csv }
    } catch {
        if ($backup) { Write-Warning "El respaldo local sigue disponible: $($backup.Json)" }
        if ($sendAttempted) { Write-Warning 'Compruebe si la fila ya existe en Excel antes de reenviar. No se reintentara automaticamente.' }
        throw
    } finally {
        if ($previousRequestContext) {
            try { $null = Set-MgRequestContext -MaxRetry $previousRequestContext.MaxRetry -ClientTimeout $previousRequestContext.ClientTimeoutSeconds } catch { Write-Verbose 'No se pudo restaurar el contexto de solicitudes.' }
        }
        if ($connected) { try { $null = Disconnect-MgGraph -ErrorAction Stop } catch { Write-Verbose 'No se pudo cerrar la sesion Graph.' } }
        [Net.ServicePointManager]::SecurityProtocol = $previousTls
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $arguments = @{} + $PSBoundParameters
    if (-not $arguments.ContainsKey('ConfigPath')) { $arguments.ConfigPath = $ConfigPath }
    if (-not $arguments.ContainsKey('DirectorioSalida')) { $arguments.DirectorioSalida = $DirectorioSalida }
    Invoke-InventoryMain @arguments
}
