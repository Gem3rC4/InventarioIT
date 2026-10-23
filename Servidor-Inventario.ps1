#requires -Version 5.1
<#
.SYNOPSIS
Receptor central de inventario para una red local y un Excel de Microsoft 365.
.DESCRIPTION
Guarda cada solicitud antes de responder. Un unico proceso escribe las filas en Excel.
Iniciar este archivo abre una sesion Microsoft en el receptor; importarlo con punto no inicia servicios.
La configuracion real, clave y cola deben permanecer fuera del repositorio.
#>
[CmdletBinding()]
param(
    [string]$ConfigPath = (Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos\servidor.config.json'),
    [string]$ListenUri,
    [string]$QueueDirectory,
    [ValidateRange(1,1000)][int]$MaxPending = 1000,
    [switch]$InstalarDependencias
)

function Get-InventoryServerValue {
    param([AllowNull()][object]$Object, [string]$Name)
    if ($null -eq $Object) { return $null }
    if ($Object -is [System.Collections.IDictionary]) { return $Object[$Name] }
    $property = $Object.PSObject.Properties[$Name]
    if ($null -ne $property) { return $property.Value }
    return $null
}

function New-InventoryServerError {
    param([string]$Message, [int]$StatusCode = 400)
    $exception = New-Object System.InvalidOperationException($Message)
    $exception.Data['HttpStatusCode'] = $StatusCode
    return $exception
}

function Assert-InventoryListenUri {
    param([Parameter(Mandatory=$true)][string]$ListenUri)
    $parsed = $null
    if (-not [uri]::TryCreate($ListenUri,[UriKind]::Absolute,[ref]$parsed) -or
        $parsed.Scheme -notin @('http','https') -or $parsed.AbsolutePath -ne '/' -or
        $parsed.Query -or $parsed.Fragment -or $parsed.UserInfo -or $parsed.Port -lt 1) {
        throw 'ListenUri debe ser http(s)://IP-privada:puerto/ con ruta raiz, sin comodines.'
    }
    $hostName = $parsed.DnsSafeHost
    $allowed = $hostName -in @('localhost','127.0.0.1','::1')
    if (-not $allowed -and $hostName -match '^([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})\.([0-9]{1,3})$') {
        $parts = @($hostName.Split('.') | ForEach-Object { [int]$_ })
        if (@($parts | Where-Object { $_ -gt 255 }).Count -eq 0) {
            $allowed = $parts[0] -eq 10 -or
                ($parts[0] -eq 172 -and $parts[1] -ge 16 -and $parts[1] -le 31) -or
                ($parts[0] -eq 192 -and $parts[1] -eq 168)
        }
    }
    if (-not $allowed) { throw 'ListenUri solo admite loopback o una direccion IPv4 privada RFC1918 de la red local.' }
    return $parsed.AbsoluteUri
}

function Test-InventoryApiKey {
    param([AllowNull()][string]$Authorization, [AllowNull()][string]$ApiKey)
    if ([string]::IsNullOrEmpty($ApiKey) -or [string]::IsNullOrEmpty($Authorization) -or
        $Authorization.Length -gt 2048 -or -not $Authorization.StartsWith('Bearer ',[StringComparison]::OrdinalIgnoreCase)) { return $false }
    $provided = [Text.Encoding]::UTF8.GetBytes($Authorization.Substring(7))
    $expected = [Text.Encoding]::UTF8.GetBytes($ApiKey)
    $difference = $provided.Length -bxor $expected.Length
    $length = [math]::Max($provided.Length,$expected.Length)
    for ($i=0; $i -lt $length; $i++) {
        $left = 0; $right = 0
        if ($i -lt $provided.Length) { $left = [int]$provided[$i] }
        if ($i -lt $expected.Length) { $right = [int]$expected[$i] }
        $difference = $difference -bor ($left -bxor $right)
    }
    return $difference -eq 0
}

function ConvertTo-InventoryRequestId {
    param([Parameter(Mandatory=$true)][string]$RequestId)
    $parsed = [guid]::Empty
    if (-not [guid]::TryParseExact($RequestId,'D',[ref]$parsed) -or $parsed -eq [guid]::Empty) {
        throw (New-InventoryServerError 'requestId debe ser un GUID valido con guiones.' 400)
    }
    return $parsed.ToString('D').ToLowerInvariant()
}

function ConvertTo-InventoryEnvelope {
    param([Parameter(Mandatory=$true)][AllowNull()][object]$Envelope)
    if ($null -eq $Envelope -or $Envelope -is [string] -or $Envelope -is [ValueType] -or $Envelope -is [array]) {
        throw (New-InventoryServerError 'El cuerpo debe ser un objeto de solicitud.' 400)
    }
    $names = @($Envelope.PSObject.Properties.Name)
    if ($Envelope -is [System.Collections.IDictionary]) { $names = @($Envelope.Keys) }
    if ($names.Count -ne 3 -or $names -cnotcontains 'schemaVersion' -or
        $names -cnotcontains 'requestId' -or $names -cnotcontains 'record') {
        throw (New-InventoryServerError 'Se requieren exactamente schemaVersion, requestId y record.' 400)
    }
    $schema = Get-InventoryServerValue $Envelope 'schemaVersion'
    if ($null -eq $schema -or $schema -is [bool] -or $schema -is [string] -or
        $schema -isnot [ValueType] -or [double]$schema -ne 1) {
        throw (New-InventoryServerError 'schemaVersion debe ser el numero 1.' 400)
    }
    $idValue = Get-InventoryServerValue $Envelope 'requestId'
    if ($idValue -isnot [string]) { throw (New-InventoryServerError 'requestId debe ser texto.' 400) }
    $requestId = ConvertTo-InventoryRequestId $idValue
    $record = Get-InventoryServerValue $Envelope 'record'
    if ($null -eq $record) { throw (New-InventoryServerError 'Falta el registro de inventario.' 400) }
    if ($record -is [System.Collections.IDictionary]) { $record = [pscustomobject]$record }
    try { Assert-InventoryRecord $record } catch { throw (New-InventoryServerError ('Registro no valido: ' + $_.Exception.Message) 400) }
    $orderedRecord = [ordered]@{}
    foreach ($field in Get-InventoryFields) {
        $value = $record.PSObject.Properties[$field].Value
        if ($field -in @('RAM_GB','DiscoTotal_GB','DiscoUsado_GB','DiscoLibre_GB') -and $null -ne $value) { $value = [double]$value }
        $orderedRecord[$field] = $value
    }
    $normalizedRecord = [pscustomobject]$orderedRecord
    $canonical = [ordered]@{ schemaVersion=1; requestId=$requestId; record=$normalizedRecord } | ConvertTo-Json -Compress -Depth 8
    $sha = [Security.Cryptography.SHA256]::Create()
    try { $hash = [BitConverter]::ToString($sha.ComputeHash([Text.Encoding]::UTF8.GetBytes($canonical))).Replace('-','').ToLowerInvariant() }
    finally { $sha.Dispose() }
    return [pscustomobject]@{ RequestId=$requestId; Record=$normalizedRecord; EnvelopeHash=$hash }
}

function Get-InventoryQueuePath {
    param([string]$QueueDirectory, [string]$RequestId)
    $id = ConvertTo-InventoryRequestId $RequestId
    $rootPath = [IO.Path]::GetFullPath($QueueDirectory)
    return Join-Path $rootPath ($id + '.json')
}

function Write-InventoryQueueItem {
    param([string]$Path, [Parameter(Mandatory=$true)][object]$Item)
    $temporary = $Path + '.' + [guid]::NewGuid().ToString('N') + '.tmp'
    try {
        $json = $Item | ConvertTo-Json -Depth 10
        [IO.File]::WriteAllText($temporary,$json,(New-Object Text.UTF8Encoding($false)))
        if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temporary,$Path,[NullString]::Value) }
        else { [IO.File]::Move($temporary,$Path) }
    } finally {
        if ([IO.File]::Exists($temporary)) { [IO.File]::Delete($temporary) }
    }
}

function Read-InventoryQueueItem {
    param([string]$Path)
    return ([IO.File]::ReadAllText($Path,[Text.Encoding]::UTF8) | ConvertFrom-Json)
}

function ConvertTo-InventoryQueueReceipt {
    param([object]$Item)
    $state = [string](Get-InventoryServerValue $Item 'status')
    $status = if ($state -eq 'inflight') { 'queued' } elseif ($state -in @('queued','sent','needs-review')) { $state } else { 'needs-review' }
    $message = switch ($status) {
        'queued' { 'Registro guardado en la cola del receptor; pendiente de confirmacion de Excel.' }
        'sent' { 'Registro confirmado en Excel.' }
        default { 'El receptor requiere revision manual. No vuelva a enviar con otro identificador.' }
    }
    $result = [ordered]@{
        requestId = [string](Get-InventoryServerValue $Item 'requestId')
        status = $status
        message = $message
        queuedAtUtc = Get-InventoryServerValue $Item 'queuedAtUtc'
        updatedAtUtc = Get-InventoryServerValue $Item 'updatedAtUtc'
    }
    if ($status -eq 'sent') { $result.rowIndex = Get-InventoryServerValue $Item 'rowIndex' }
    return [pscustomobject]$result
}

function Get-InventoryQueueReceipt {
    param([string]$QueueDirectory, [string]$RequestId)
    $path = Get-InventoryQueuePath $QueueDirectory $RequestId
    if (-not [IO.File]::Exists($path)) { throw (New-InventoryServerError 'Solicitud no encontrada.' 404) }
    return ConvertTo-InventoryQueueReceipt (Read-InventoryQueueItem $path)
}

function Add-InventoryQueueRequest {
    param([object]$Envelope, [string]$QueueDirectory, [ValidateRange(1,1000)][int]$MaxPending=1000)
    $normalized = ConvertTo-InventoryEnvelope $Envelope
    $path = Get-InventoryQueuePath $QueueDirectory $normalized.RequestId
    if ([IO.File]::Exists($path)) {
        $existing = Read-InventoryQueueItem $path
        if ([string](Get-InventoryServerValue $existing 'envelopeHash') -cne $normalized.EnvelopeHash) {
            throw (New-InventoryServerError 'requestId ya existe con un contenido diferente.' 409)
        }
        return ConvertTo-InventoryQueueReceipt $existing
    }
    $null = New-Item -ItemType Directory -Path $QueueDirectory -Force -ErrorAction Stop
    $pending = 0
    foreach ($file in @(Get-ChildItem -LiteralPath $QueueDirectory -Filter '*.json' -File -ErrorAction Stop)) {
        if ($file.BaseName -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { continue }
        try { if ((Get-InventoryServerValue (Read-InventoryQueueItem $file.FullName) 'status') -ne 'sent') { $pending++ } }
        catch { $pending++ }
    }
    if ($pending -ge $MaxPending) { throw (New-InventoryServerError 'La cola de pendientes esta llena. Reintente mas tarde con el mismo requestId.' 503) }
    $now = [DateTime]::UtcNow.ToString('o')
    $item = [pscustomobject][ordered]@{
        schemaVersion=1; requestId=$normalized.RequestId; status='queued'; envelopeHash=$normalized.EnvelopeHash
        queuedAtUtc=$now; updatedAtUtc=$now; attemptedAtUtc=$null; sentAtUtc=$null
        rowIndex=$null; lastError=$null; record=$normalized.Record
    }
    Write-InventoryQueueItem $path $item
    return ConvertTo-InventoryQueueReceipt $item
}

function Repair-InventoryQueue {
    param([string]$QueueDirectory)
    if (-not (Test-Path -LiteralPath $QueueDirectory -PathType Container)) { return }
    foreach ($file in @(Get-ChildItem -LiteralPath $QueueDirectory -Filter '*.json' -File -ErrorAction Stop)) {
        if ($file.BaseName -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { continue }
        $item = Read-InventoryQueueItem $file.FullName
        if ((Get-InventoryServerValue $item 'status') -eq 'inflight') {
            $item.status = 'needs-review'
            $item.updatedAtUtc = [DateTime]::UtcNow.ToString('o')
            $item.lastError = 'El receptor se detuvo durante un envio. Verifique Excel antes de cualquier accion manual.'
            Write-InventoryQueueItem $file.FullName $item
        }
    }
}

function Invoke-InventoryQueueWorker {
    param([string]$QueueDirectory, [Parameter(Mandatory=$true)][object]$Destination, [scriptblock]$Sender, [scriptblock]$DestinationResolver)
    $selected = $null; $item = $null
    foreach ($file in @(Get-ChildItem -LiteralPath $QueueDirectory -Filter '*.json' -File -ErrorAction Stop | Sort-Object CreationTimeUtc,Name)) {
        if ($file.BaseName -notmatch '^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$') { continue }
        try {
            $candidate = Read-InventoryQueueItem $file.FullName
            if ((Get-InventoryServerValue $candidate 'status') -eq 'queued') { $selected=$file; $item=$candidate; break }
        } catch { Write-Warning ('No se pudo leer un archivo de la cola: ' + $file.Name) }
    }
    if ($null -eq $selected) { return $null }
    $attempted = $false
    try {
        $envelope = [pscustomobject]@{ schemaVersion=1; requestId=$item.requestId; record=$item.record }
        $checked = ConvertTo-InventoryEnvelope $envelope
        if ($checked.RequestId -cne $selected.BaseName -or $checked.EnvelopeHash -cne $item.envelopeHash) { throw 'El contenido de la cola no coincide con su recibo original.' }
        if ($null -ne $DestinationResolver) { $Destination = & $DestinationResolver }
        $body = New-InventoryRowBody -Record $checked.Record -Columns $Destination.Columns
        if (-not [string]$Destination.TableUri -or -not ([string]$Destination.TableUri).StartsWith('https://graph.microsoft.com/v1.0/',[StringComparison]::OrdinalIgnoreCase)) { throw 'El destino de Graph no es valido.' }
        $item.status='inflight'
        $item.updatedAtUtc=[DateTime]::UtcNow.ToString('o')
        $item.attemptedAtUtc=$item.updatedAtUtc
        Write-InventoryQueueItem $selected.FullName $item
        $attempted=$true
        if ($null -ne $Sender) { $response = & $Sender ($Destination.TableUri + '/rows/add') $body }
        else { $response = Invoke-MgGraphRequest -Method POST -Uri ($Destination.TableUri + '/rows/add') -Body $body -ContentType 'application/json; charset=utf-8' -OutputType PSObject -ErrorAction Stop }
        $index = Get-InventoryServerValue $response 'index'
        if ($null -eq $index -or $index -is [bool] -or $index -is [string] -or $index -isnot [ValueType] -or [double]$index -lt 0 -or [double]::IsNaN([double]$index) -or [double]::IsInfinity([double]$index) -or [double]$index -ne [math]::Floor([double]$index)) { throw 'Excel no devolvio confirmacion de la fila.' }
        $item.status='sent'; $item.rowIndex=$index; $item.sentAtUtc=[DateTime]::UtcNow.ToString('o')
        $item.updatedAtUtc=$item.sentAtUtc; $item.lastError=$null
        Write-InventoryQueueItem $selected.FullName $item
    } catch {
        $item.status='needs-review'; $item.updatedAtUtc=[DateTime]::UtcNow.ToString('o')
        $prefix = if ($attempted) { 'Resultado del envio desconocido; revise Excel. ' } else { 'No se intento insertar en Excel. ' }
        $item.lastError = $prefix + $_.Exception.Message
        try { Write-InventoryQueueItem $selected.FullName $item }
        catch { Write-Warning ('No se pudo persistir el estado de revision de la solicitud ' + $selected.BaseName + '.') }
        Write-Warning ('Solicitud ' + $selected.BaseName + ': ' + $prefix)
    }
    return ConvertTo-InventoryQueueReceipt $item
}

function Read-InventoryRequestBytes {
    param([Parameter(Mandatory=$true)][IO.Stream]$InputStream, [long]$ContentLength=-1, [ValidateRange(1,65536)][int]$MaxBytes=65536)
    if ($ContentLength -gt $MaxBytes) { throw (New-InventoryServerError 'El cuerpo excede el limite de 64 KiB.' 413) }
    $memory = New-Object IO.MemoryStream
    try {
        $buffer = [byte[]]::new(4096)
        do {
            $count = $InputStream.Read($buffer,0,$buffer.Length)
            if ($memory.Length + $count -gt $MaxBytes) { throw (New-InventoryServerError 'El cuerpo excede el limite de 64 KiB.' 413) }
            if ($count -gt 0) { $memory.Write($buffer,0,$count) }
        } while ($count -gt 0)
        return ,$memory.ToArray()
    } finally { $memory.Dispose() }
}

function Invoke-InventoryApi {
    param([string]$Method,[string]$Path,[string]$Authorization,[AllowNull()][byte[]]$BodyBytes,[string]$ContentType,
        [string]$ApiKey,[string]$QueueDirectory,[ValidateRange(1,1000)][int]$MaxPending=1000)
    try {
        if (-not (Test-InventoryApiKey $Authorization $ApiKey)) { throw (New-InventoryServerError 'No autorizado.' 401) }
        if ($null -ne $BodyBytes -and $BodyBytes.Length -gt 65536) { throw (New-InventoryServerError 'El cuerpo excede el limite de 64 KiB.' 413) }
        if ($Path -ceq '/inventario') {
            if ($Method -cne 'POST') { throw (New-InventoryServerError 'Metodo no permitido.' 405) }
            if ($ContentType -notmatch '^application/json(?:\s*;\s*charset\s*=\s*"?utf-8"?)?\s*$') { throw (New-InventoryServerError 'Se requiere application/json con codificacion UTF-8.' 415) }
            if ($null -eq $BodyBytes -or $BodyBytes.Length -eq 0) { throw (New-InventoryServerError 'El cuerpo JSON esta vacio.' 400) }
            try {
                $utf8 = New-Object Text.UTF8Encoding($false,$true)
                $text = $utf8.GetString($BodyBytes)
                $envelope = $text.TrimStart([char]0xFEFF) | ConvertFrom-Json -ErrorAction Stop
            } catch { throw (New-InventoryServerError 'El cuerpo no es JSON UTF-8 valido.' 400) }
            $receipt = Add-InventoryQueueRequest -Envelope $envelope -QueueDirectory $QueueDirectory -MaxPending $MaxPending
            $code = if ($receipt.status -eq 'queued') { 202 } else { 200 }
            return [pscustomobject]@{ StatusCode=$code; Body=$receipt }
        }
        if ($Path -cmatch '^/estado/([^/]+)$') {
            if ($Method -cne 'GET') { throw (New-InventoryServerError 'Metodo no permitido.' 405) }
            $receipt = Get-InventoryQueueReceipt -QueueDirectory $QueueDirectory -RequestId $Matches[1]
            return [pscustomobject]@{ StatusCode=200; Body=$receipt }
        }
        throw (New-InventoryServerError 'Ruta no encontrada.' 404)
    } catch {
        $code = Get-InventoryServerValue $_.Exception.Data 'HttpStatusCode'
        if ($null -eq $code) { $code=500; $message='No se pudo procesar la solicitud. El receptor requiere revision.' }
        else { $message=$_.Exception.Message }
        return [pscustomobject]@{ StatusCode=[int]$code; Body=[pscustomobject]@{ status='error'; message=$message } }
    }
}

function Send-InventoryHttpResponse {
    param([object]$Context,[object]$Result)
    $bytes = [Text.Encoding]::UTF8.GetBytes(($Result.Body | ConvertTo-Json -Compress -Depth 8))
    try {
        $Context.Response.StatusCode=$Result.StatusCode
        $Context.Response.ContentType='application/json; charset=utf-8'
        $Context.Response.ContentEncoding=[Text.Encoding]::UTF8
        $Context.Response.ContentLength64=$bytes.Length
        $Context.Response.KeepAlive=$false
        $Context.Response.Headers['Cache-Control']='no-store'
        $Context.Response.OutputStream.Write($bytes,0,$bytes.Length)
    } finally { $Context.Response.Close() }
}

function Receive-InventoryHttpRequest {
    param([object]$Context,[string]$ApiKey,[string]$QueueDirectory,[int]$MaxPending=1000)
    try {
        $request=$Context.Request
        if (-not (Test-InventoryApiKey $request.Headers['Authorization'] $ApiKey)) {
            $result=[pscustomobject]@{ StatusCode=401; Body=[pscustomobject]@{ status='error'; message='No autorizado.' } }
        } else {
            $bytes = Read-InventoryRequestBytes -InputStream $request.InputStream -ContentLength $request.ContentLength64
            $result=Invoke-InventoryApi -Method $request.HttpMethod -Path $request.Url.AbsolutePath -Authorization $request.Headers['Authorization'] -BodyBytes $bytes -ContentType $request.ContentType -ApiKey $ApiKey -QueueDirectory $QueueDirectory -MaxPending $MaxPending
        }
    } catch {
        $code=Get-InventoryServerValue $_.Exception.Data 'HttpStatusCode'
        if ($null -eq $code) { $code=500; $message='No se pudo guardar la solicitud.' } else { $message=$_.Exception.Message }
        $result=[pscustomobject]@{ StatusCode=[int]$code; Body=[pscustomobject]@{ status='error'; message=$message } }
    }
    try { Send-InventoryHttpResponse -Context $Context -Result $result }
    catch { Write-Warning 'No se pudo devolver la respuesta HTTP; los recibos ya persistidos se conservan.' }
}

function Start-InventoryServer {
    [CmdletBinding()]
    param([string]$ConfigPath,[string]$ListenUri,[string]$QueueDirectory,[ValidateRange(1,1000)][int]$MaxPending=1000,
        [switch]$InstalarDependencias,[string]$HelperModulePath)
    Set-StrictMode -Version 2.0
    $ErrorActionPreference='Stop'
    if ($env:OS -ne 'Windows_NT') { throw 'El receptor debe ejecutarse en Windows PowerShell 5.1 o posterior.' }
    if (-not $HelperModulePath) { $HelperModulePath=Join-Path $PSScriptRoot 'Inventario.Funciones.psm1' }
    Import-Module $HelperModulePath -Force -ErrorAction Stop
    $config=Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 -ErrorAction Stop | ConvertFrom-Json
    if (-not $ListenUri) { $ListenUri=[string](Get-InventoryServerValue $config 'ListenUri') }
    if (-not $ListenUri) { $ListenUri='http://127.0.0.1:8765/' }
    $ListenUri=Assert-InventoryListenUri $ListenUri
    if (-not $QueueDirectory) { $QueueDirectory=[string](Get-InventoryServerValue $config 'QueueDirectory') }
    if (-not $QueueDirectory) { $QueueDirectory=Join-Path ([Environment]::GetFolderPath('LocalApplicationData')) 'InventarioEquipos\cola' }
    $QueueDirectory=[IO.Path]::GetFullPath($QueueDirectory)
    $apiKey=[string](Get-InventoryServerValue $config 'ApiKey')
    if (-not $apiKey) { $apiKey=[string]$env:INVENTARIO_API_KEY }
    if ($apiKey.Length -lt 32 -or $apiKey.Length -gt 512 -or $apiKey -match '[\s\x00-\x1f\x7f]') { throw 'Configure una ApiKey aleatoria de al menos 32 caracteres en el archivo local o INVENTARIO_API_KEY.' }
    $excelUrl=[string](Get-InventoryServerValue $config 'ExcelUrl')
    if (-not $excelUrl) { throw 'Falta ExcelUrl en la configuracion del receptor.' }
    $null=ConvertTo-SharingToken $excelUrl
    $null=New-Item -ItemType Directory -Path $QueueDirectory -Force
    # Un solo receptor puede utilizar este directorio. El bloqueo se libera al terminar.
    $lockPath=Join-Path $QueueDirectory 'receptor.lock'
    $queueLock=$null; $listener=$null; $previousContext=$null
    $previousTls=[Net.ServicePointManager]::SecurityProtocol
    try {
        $queueLock=[IO.File]::Open($lockPath,[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
        Repair-InventoryQueue $QueueDirectory
        [Net.ServicePointManager]::SecurityProtocol=$previousTls -bor [Net.SecurityProtocolType]::Tls12
        Initialize-GraphModule -Install:$InstalarDependencias
        Import-Module Microsoft.Graph.Authentication -MinimumVersion 2.0.0 -ErrorAction Stop
        $connect=@{ Scopes=@('Files.ReadWrite'); ContextScope='CurrentUser'; UseDeviceAuthentication=$true; NoWelcome=$true; ErrorAction='Stop' }
        $tenant=[string](Get-InventoryServerValue $config 'TenantId'); $client=[string](Get-InventoryServerValue $config 'ClientId')
        if ($tenant) { $connect.TenantId=$tenant }; if ($client) { $connect.ClientId=$client }
        Write-Host 'Autentique SOLO el receptor con una cuenta que pueda editar el Excel.' -ForegroundColor Cyan
        $null=Connect-MgGraph @connect
        $graphContext=Get-MgContext
        if ((Get-InventoryServerValue $graphContext 'AuthType') -ne 'Delegated') { throw 'Se requiere una sesion delegada de Microsoft Graph.' }
        $saved=Get-MgRequestContext
        $previousContext=@{ MaxRetry=[int]$saved.MaxRetry; ClientTimeout=[int]$saved.ClientTimeout.TotalSeconds }
        $null=Set-MgRequestContext -MaxRetry 0 -ClientTimeout 120
        $destination=Get-ExcelDestination -Url $excelUrl -TableName ([string](Get-InventoryServerValue $config 'Tabla'))
        $resolveDestination = { Get-ExcelDestination -Url $excelUrl -TableName ([string](Get-InventoryServerValue $config 'Tabla')) }
        $listener=New-Object Net.HttpListener
        $listener.TimeoutManager.EntityBody=[timespan]::FromSeconds(15)
        $listener.TimeoutManager.HeaderWait=[timespan]::FromSeconds(15)
        $listener.Prefixes.Add($ListenUri)
        $listener.Start()
        Write-Host ('Receptor activo: ' + $ListenUri + ' | Tabla: ' + $destination.TableName) -ForegroundColor Green
        Write-Host ('Cola durable: ' + $QueueDirectory)
        Write-Host 'Use Ctrl+C para detener. Los recibos pendientes permanecen en disco.'
        $pendingContext=$listener.GetContextAsync()
        while ($listener.IsListening) {
            if ($pendingContext.Wait(250)) {
                $context=$pendingContext.GetAwaiter().GetResult()
                Receive-InventoryHttpRequest -Context $context -ApiKey $apiKey -QueueDirectory $QueueDirectory -MaxPending $MaxPending
                $pendingContext=$listener.GetContextAsync()
            }
            try {
                $workerResult=Invoke-InventoryQueueWorker -QueueDirectory $QueueDirectory -Destination $destination -DestinationResolver $resolveDestination
                if ($null -ne $workerResult) {
                    $color = if ($workerResult.status -eq 'sent') { 'Green' } else { 'Yellow' }
                    Write-Host ('[{0}] Solicitud {1}' -f $workerResult.status,$workerResult.requestId) -ForegroundColor $color
                }
            }
            catch { Write-Warning 'El trabajador no pudo procesar la cola; el receptor sigue atendiendo solicitudes.' }
        }
    } finally {
        if ($null -ne $listener) { try { $listener.Stop(); $listener.Close() } catch {} }
        if ($previousContext) { try { $null=Set-MgRequestContext -MaxRetry $previousContext.MaxRetry -ClientTimeout $previousContext.ClientTimeout } catch {} }
        if ($null -ne $queueLock) { $queueLock.Dispose() }
        [Net.ServicePointManager]::SecurityProtocol=$previousTls
        # CurrentUser conserva la cache protegida de MSAL en el receptor. No se exportan tokens.
        # Para retirar la sesion del receptor, use Disconnect-MgGraph de forma explicita.
    }
}

if ($MyInvocation.InvocationName -ne '.') {
    $serverArguments=@{} + $PSBoundParameters
    if (-not $serverArguments.ContainsKey('ConfigPath')) { $serverArguments.ConfigPath=$ConfigPath }
    $serverArguments.HelperModulePath=Join-Path $PSScriptRoot 'Inventario.Funciones.psm1'
    Start-InventoryServer @serverArguments
}
