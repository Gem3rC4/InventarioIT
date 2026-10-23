#requires -Version 5.1
<#
.SYNOPSIS
Prueba el receptor y su cola en disco sin iniciar listener, Microsoft Graph ni Excel.
#>
[CmdletBinding()]
param([string]$ServerPath,[string]$TestRoot=([IO.Path]::GetTempPath()))
$ErrorActionPreference='Stop'
if (-not $ServerPath) { $ServerPath=Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) 'Servidor-Inventario.ps1' }
Import-Module (Join-Path (Split-Path $ServerPath -Parent) 'Inventario.Funciones.psm1') -Force -DisableNameChecking
. $ServerPath
$script:Passed=0
$script:Failures=New-Object 'System.Collections.Generic.List[string]'
$script:TempRoot=Join-Path $TestRoot ('it-r-'+[guid]::NewGuid().ToString('N'))
$null=New-Item -ItemType Directory -Path $script:TempRoot
$script:TestKey='clave-sintetica-pruebas-sin-uso-real-123456'
$script:SendCount=0
function Assert-True { param([bool]$Condition,[string]$Message); if (-not $Condition) { throw $Message } }
function Assert-Equal {
    param([AllowNull()][object]$Actual,[AllowNull()][object]$Expected,[string]$Message)
    if ($null -eq $Expected) { if ($null -ne $Actual) { throw "$Message (se esperaba null)" } }
    elseif ($Actual -cne $Expected) { throw "$Message (esperado: $Expected; recibido: $Actual)" }
}
function Assert-Throws {
    param([scriptblock]$Action,[string]$Message,[int]$Status=0)
    $thrown=$false
    try { & $Action|Out-Null } catch {
        $thrown=$true
        if ($Status) { Assert-Equal ([int]$_.Exception.Data['HttpStatusCode']) $Status 'Codigo de error incorrecto.' }
    }
    Assert-True $thrown $Message
}
function Invoke-Test {
    param([string]$Name,[scriptblock]$Action)
    try { & $Action; $script:Passed++; Write-Host "OK: $Name" -ForegroundColor Green }
    catch { $script:Failures.Add(("{0}: {1}" -f $Name,$_.Exception.Message)); Write-Host ("FALLO: {0}: {1}" -f $Name,$_.Exception.Message) -ForegroundColor Red }
}
function New-TestRecord {
    $record=[ordered]@{}
    foreach ($field in Get-InventoryFields) { $record[$field]='valor-'+$field }
    $record.Equipo='PC-PRUEBA'; $record.Area='Tecnolog'+[char]0x00ED+'a'
    $record.RAM_GB=16.25; $record.DiscoTotal_GB=256; $record.DiscoUsado_GB=64.5; $record.DiscoLibre_GB=$null
    [pscustomobject]$record
}
function New-TestEnvelope { [pscustomobject][ordered]@{ schemaVersion=1; requestId=[guid]::NewGuid().ToString('D'); record=(New-TestRecord) } }
function New-TestDirectory {
    $path=Join-Path $script:TempRoot ([guid]::NewGuid().ToString('N'))
    $null=New-Item -ItemType Directory -Path $path
    $path
}
function Invoke-TestApi {
    param([object]$Envelope,[string]$Directory,[string]$Method='POST',[string]$Path='/inventario',[string]$Authorization=('Bearer '+$script:TestKey),[string]$ContentType='application/json; charset=utf-8',[byte[]]$Bytes,[int]$MaxPending=1000)
    if (-not $PSBoundParameters.ContainsKey('Bytes')) {
        $Bytes=[Text.Encoding]::UTF8.GetBytes(($Envelope|ConvertTo-Json -Depth 8 -Compress))
    }
    Invoke-InventoryApi -Method $Method -Path $Path -Authorization $Authorization -BodyBytes $Bytes -ContentType $ContentType -ApiKey $script:TestKey -QueueDirectory $Directory -MaxPending $MaxPending
}
function New-TestDestination {
    $columns=@(); $index=0
    foreach ($name in Get-InventoryFields) { $columns += [pscustomobject]@{name=$name;index=$index}; $index++ }
    [pscustomobject]@{ TableUri='https://graph.microsoft.com/v1.0/drives/TEST/items/TEST/workbook/tables/TEST'; TableName='Inventario'; Columns=$columns }
}
# Si el codigo intenta enviar de verdad, esta funcion hace fallar la prueba.
function Invoke-MgGraphRequest { throw 'La suite no permite llamadas reales a Microsoft Graph.' }
function Invoke-InventarioGraphRequest { throw 'La suite no permite llamadas reales a Microsoft Graph.' }
try {
    Invoke-Test 'Clave Bearer correcta y rechaza faltantes, prefijos o clave distinta' {
        Assert-True (Test-InventoryApiKey ('Bearer '+$script:TestKey) $script:TestKey) 'Rechazo clave correcta.'
        foreach ($header in @($null,'','Basic '+$script:TestKey,'Bearer diferente','Bearer '+$script:TestKey+'x')) {
            Assert-True (-not (Test-InventoryApiKey $header $script:TestKey)) 'Acepto clave incorrecta.'
        }
    }
    Invoke-Test 'Escucha explicita privada y bloquea comodines/publico/rutas' {
        foreach ($uri in @('http://127.0.0.1:8765/','http://192.168.1.10:8765/','https://10.1.1.1:8765/')) { $null=Assert-InventoryListenUri $uri }
        foreach ($uri in @('http://+:8765/','http://*:8765/','http://0.0.0.0:8765/','http://8.8.8.8:8765/','http://192.168.1.10:8765/otra/')) {
            Assert-Throws { Assert-InventoryListenUri $uri } 'Acepto direccion de escucha indebida.'
        }
    }
    Invoke-Test 'Envelope normaliza ID, orden y numeros para hash estable' {
        $envelope=New-TestEnvelope
        $envelope.requestId=$envelope.requestId.ToUpperInvariant()
        $a=ConvertTo-InventoryEnvelope $envelope
        Assert-Equal $a.RequestId $envelope.requestId.ToLowerInvariant() 'ID no canonico.'
        Assert-True ($a.EnvelopeHash -cmatch '^[0-9a-f]{64}$') 'Hash SHA256 incorrecto.'
        $reverse=[ordered]@{}; $fields=@(Get-InventoryFields); [array]::Reverse($fields)
        foreach ($field in $fields) { $reverse[$field]=$envelope.record.PSObject.Properties[$field].Value }
        $envelope.record=[pscustomobject]$reverse
        $b=ConvertTo-InventoryEnvelope $envelope
        Assert-Equal $a.EnvelopeHash $b.EnvelopeHash 'El orden de propiedades cambio el hash.'
        Assert-Equal $a.Record.Area $b.Record.Area 'Normalizacion cambio Unicode.'
        Assert-Equal $a.Record.DiscoLibre_GB $null 'Normalizacion perdio null.'
    }
    Invoke-Test 'Rechaza version booleana/texto, GUID vacio/traversal y campos extra' {
        foreach ($version in @($true,'1',2,$null)) {
            $envelope=New-TestEnvelope; $envelope.schemaVersion=$version
            Assert-Throws { ConvertTo-InventoryEnvelope $envelope } 'Acepto version incorrecta.' 400
        }
        foreach ($id in @('..\archivo',[guid]::Empty.ToString('D'),'no-guid')) {
            $envelope=New-TestEnvelope; $envelope.requestId=$id
            Assert-Throws { ConvertTo-InventoryEnvelope $envelope } 'Acepto ID peligroso o vacio.' 400
        }
        $envelope=New-TestEnvelope; $envelope|Add-Member extra 'x'
        Assert-Throws { ConvertTo-InventoryEnvelope $envelope } 'Acepto propiedad extra.' 400
        $envelope=New-TestEnvelope; $envelope.record.PSObject.Properties.Remove('Area')
        Assert-Throws { ConvertTo-InventoryEnvelope $envelope } 'Acepto registro incompleto.' 400
    }
    Invoke-Test 'API autentica antes de leer y valida metodo/ruta/content-type/JSON' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        Assert-Equal (Invoke-TestApi $envelope $directory -Authorization '').StatusCode 401 'Permite acceso anonimo.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -File).Count 0 'Creo cola sin autenticar.'
        Assert-Equal (Invoke-TestApi $envelope $directory -Method GET).StatusCode 405 'Acepto metodo incorrecto.'
        Assert-Equal (Invoke-TestApi $envelope $directory -Path '/otra').StatusCode 404 'Acepto ruta incorrecta.'
        Assert-Equal (Invoke-TestApi $envelope $directory -ContentType 'text/plain').StatusCode 415 'Acepto formato incorrecto.'
        Assert-Equal (Invoke-TestApi $envelope $directory -Bytes ([Text.Encoding]::UTF8.GetBytes('{'))).StatusCode 400 'Acepto JSON roto.'
        Assert-Equal (Invoke-TestApi $envelope $directory -Bytes ([byte[]]@(0xC3,0x28))).StatusCode 400 'Acepto UTF-8 invalido.'
        Assert-Equal (Invoke-TestApi $envelope $directory -Bytes ([byte[]]@())).StatusCode 400 'Acepto cuerpo vacio.'
    }
    Invoke-Test 'POST persiste antes de202 y GET solo devuelve recibo publico' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $result=Invoke-TestApi $envelope $directory
        Assert-Equal $result.StatusCode 202 'Una nueva solicitud no quedo en cola.'
        Assert-Equal $result.Body.status 'queued' 'Presento recibido como Excel confirmado.'
        $files=@(Get-ChildItem -LiteralPath $directory -Filter '*.json' -File)
        Assert-Equal $files.Count 1 'La solicitud no es durable.'
        Assert-Equal $files[0].BaseName $envelope.requestId 'El nombre no es el ID validado.'
        $saved=Get-Content -LiteralPath $files[0].FullName -Raw -Encoding UTF8|ConvertFrom-Json
        Assert-Equal $saved.record.Area $envelope.record.Area 'El disco perdio Unicode.'
        Assert-Equal $saved.record.DiscoLibre_GB $null 'El disco perdio null.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.tmp').Count 0 'Dejo archivo temporal.'
        $receipt=Invoke-TestApi $null $directory -Method GET -Path ('/estado/'+$envelope.requestId) -Bytes ([byte[]]@())
        Assert-Equal $receipt.StatusCode 200 'No encuentra recibo persistido.'
        Assert-Equal $receipt.Body.status 'queued' 'Estado de cola incorrecto.'
        foreach ($privateName in @('record','envelopeHash','ApiKey','lastError')) {
            Assert-True ($null -eq $receipt.Body.PSObject.Properties[$privateName]) 'Recibo expone datos privados.'
        }
        Assert-Equal (Invoke-TestApi $null $directory -Method GET -Path ('/estado/'+[guid]::NewGuid().ToString('D')) -Bytes ([byte[]]@())).StatusCode 404 'Invento recibo inexistente.'
    }
    Invoke-Test 'Deduplica mismo ID y409 si cambia el contenido' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Invoke-TestApi $envelope $directory
        Assert-Equal (Invoke-TestApi $envelope $directory).Body.status 'queued' 'Duplicado cambio estado.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.json').Count 1 'Duplico archivo del ID.'
        $envelope.record.Area='OTRA'
        Assert-Equal (Invoke-TestApi $envelope $directory).StatusCode 409 'Sobrescribio ID con contenido distinto.'
        $saved=Get-Content -LiteralPath (Get-InventoryQueuePath $directory $envelope.requestId) -Raw -Encoding UTF8|ConvertFrom-Json
        Assert-True ($saved.record.Area -ne 'OTRA') 'Modifico contenido original.'
    }
    Invoke-Test 'Limite real de64KiB incluye UTF-8 y ContentLength desconocido o falso' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $envelope.record.Area=([string][char]0x00F1)*40000
        $bytes=[Text.Encoding]::UTF8.GetBytes(($envelope|ConvertTo-Json -Depth 6 -Compress))
        Assert-True ($bytes.Length -gt 65536) 'El fixture no supera limite por bytes.'
        Assert-Equal (Invoke-TestApi $null $directory -Bytes $bytes).StatusCode 413 'Acepto mas de64KiB.'
        $exact=[byte[]]::new(65536); $over=[byte[]]::new(65537)
        foreach ($length in @(-1,1)) {
            $stream=New-Object IO.MemoryStream(,$over)
            try { Assert-Throws { Read-InventoryRequestBytes $stream $length } 'Confio en ContentLength falso.' 413 } finally { $stream.Dispose() }
        }
        $stream=New-Object IO.MemoryStream(,$exact)
        try { Assert-Equal (Read-InventoryRequestBytes $stream -1).Length 65536 'Rechazo limite exacto.' } finally { $stream.Dispose() }
    }
    Invoke-Test 'Cola llena503 conserva ID existente y no agrega otro' {
        $directory=New-TestDirectory; $a=New-TestEnvelope; $b=New-TestEnvelope
        Assert-Equal (Invoke-TestApi $a $directory -MaxPending 1).StatusCode 202 'No acepto primer registro.'
        Assert-Equal (Invoke-TestApi $b $directory -MaxPending 1).StatusCode 503 'Excedio capacidad de pendientes.'
        Assert-Equal (Invoke-TestApi $a $directory -MaxPending 1).StatusCode 202 'Rechazo duplicado ya persistido.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.json').Count 1 'La cola excedio capacidad.'
    }
    Invoke-Test 'Worker confirma indice0, matriz18 y no vuelve a enviar sent' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Add-InventoryQueueRequest $envelope $directory
        $script:SendCount=0
        $sender={
            param($Uri,$Body)
            $script:SendCount++
            Assert-True ($Uri.EndsWith('/rows/add')) 'Ruta Graph incorrecta.'
            Assert-True ($Body -match '"values":\[\[') 'Matriz JSON incorrecta en PS5.1.'
            $decoded=$Body|ConvertFrom-Json
            Assert-Equal $decoded.values[0].Count 18 'Matriz perdio valores.'
            Assert-Equal $decoded.values[0][7] 16.25 'Cambio valor numerico.'
            Assert-Equal $decoded.values[0][10] $null 'Omitio null.'
            $path=Get-InventoryQueuePath $directory $envelope.requestId
            Assert-Equal (Read-InventoryQueueItem $path).status 'inflight' 'No persistio inflight antes de POST.'
            [pscustomobject]@{index=0}
        }
        $receipt=Invoke-InventoryQueueWorker $directory (New-TestDestination) $sender
        Assert-Equal $receipt.status 'sent' 'No confirmo indice cero.'
        Assert-Equal $receipt.rowIndex 0 'Indice cero incorrecto.'
        Assert-Equal (Get-InventoryQueueReceipt $directory $envelope.requestId).status 'sent' 'Estado enviado no durable.'
        $null=Invoke-InventoryQueueWorker $directory (New-TestDestination) $sender
        Assert-Equal $script:SendCount 1 'Volvio a enviar sent.'
        Assert-Equal (Invoke-TestApi $envelope $directory).Body.status 'sent' 'Duplicado perdio confirmacion.'
    }
    Invoke-Test 'POST ambiguo queda needs-review y ningun ciclo lo reenvia' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Add-InventoryQueueRequest $envelope $directory
        $script:SendCount=0
        $sender={param($Uri,$Body); $script:SendCount++; throw 'Timeout sintetico despues de POST.'}
        $receipt=Invoke-InventoryQueueWorker $directory (New-TestDestination) $sender -WarningAction SilentlyContinue
        Assert-Equal $receipt.status 'needs-review' 'Timeout no quedo para revision.'
        $null=Invoke-InventoryQueueWorker $directory (New-TestDestination) $sender
        Assert-Equal $script:SendCount 1 'Reintento POST ambiguo.'
        Repair-InventoryQueue $directory
        Assert-Equal (Invoke-TestApi $envelope $directory).Body.status 'needs-review' 'Duplicado reencolo revision.'
    }
    Invoke-Test 'Confirmacion invalida no marca sent' {
        foreach ($index in @($null,$true,'0',-1,0.5,[double]::NaN,[double]::PositiveInfinity)) {
            $directory=New-TestDirectory; $envelope=New-TestEnvelope
            $null=Add-InventoryQueueRequest $envelope $directory
            $script:BadIndex=$index; $script:SendCount=0
            $sender={param($Uri,$Body); $script:SendCount++; [pscustomobject]@{index=$script:BadIndex}}
            $receipt=Invoke-InventoryQueueWorker $directory (New-TestDestination) $sender -WarningAction SilentlyContinue
            Assert-Equal $receipt.status 'needs-review' ('Acepto indice no valido: '+$index)
            $null=Invoke-InventoryQueueWorker $directory (New-TestDestination) $sender
            Assert-Equal $script:SendCount 1 'Reintento confirmacion invalida.'
        }
    }
    Invoke-Test 'Recuperacion de inflight requiere revision y no modifica queued o sent' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Add-InventoryQueueRequest $envelope $directory
        $path=Get-InventoryQueuePath $directory $envelope.requestId
        $item=Read-InventoryQueueItem $path; $item.status='inflight'; Write-InventoryQueueItem $path $item
        Repair-InventoryQueue $directory
        Assert-Equal (Get-InventoryQueueReceipt $directory $envelope.requestId).status 'needs-review' 'Reinicio reencolo envio ambiguo.'
        $script:SendCount=0
        $null=Invoke-InventoryQueueWorker $directory (New-TestDestination) { $script:SendCount++; [pscustomobject]@{index=0} }
        Assert-Equal $script:SendCount 0 'Reinicio reintento envio ambiguo.'
        $queued=New-TestEnvelope; $null=Add-InventoryQueueRequest $queued $directory
        Repair-InventoryQueue $directory
        Assert-Equal (Get-InventoryQueueReceipt $directory $queued.requestId).status 'queued' 'Reinicio modifico registro sin intento.'
    }
    Invoke-Test 'Cola alterada se detecta antes de enviar y queda en revision' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Add-InventoryQueueRequest $envelope $directory
        $path=Get-InventoryQueuePath $directory $envelope.requestId
        $item=Read-InventoryQueueItem $path; $item.record.Area='ALTERADA'; Write-InventoryQueueItem $path $item
        $script:SendCount=0
        $receipt=Invoke-InventoryQueueWorker $directory (New-TestDestination) { $script:SendCount++; [pscustomobject]@{index=0} } -WarningAction SilentlyContinue
        Assert-Equal $receipt.status 'needs-review' 'No detecto cola alterada.'
        Assert-Equal $script:SendCount 0 'Envio datos alterados.'
        Assert-True ((Read-InventoryQueueItem $path).lastError.StartsWith('No se intento')) 'Confundio fallo previo con envio ambiguo.'
    }
    Invoke-Test 'Refresca columnas por registro y evita leer destino sin pendientes' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Add-InventoryQueueRequest $envelope $directory
        $updated=New-TestDestination; $names=@(Get-InventoryFields); [array]::Reverse($names)
        for ($i=0; $i -lt $names.Count; $i++) { $updated.Columns[$i].name=$names[$i] }
        $script:UpdatedDestination=$updated; $script:ResolveCount=0; $script:SendCount=0
        $resolver={ $script:ResolveCount++; $script:UpdatedDestination }
        $sender={
            param($Uri,$Body)
            $script:SendCount++
            $decoded=$Body|ConvertFrom-Json
            Assert-Equal $decoded.values[0][0] (New-TestRecord).Area 'No adapto los valores al nuevo orden.'
            [pscustomobject]@{index=0}
        }
        $receipt=Invoke-InventoryQueueWorker -QueueDirectory $directory -Destination (New-TestDestination) -Sender $sender -DestinationResolver $resolver
        Assert-Equal $receipt.status 'sent' 'Fallo usando nuevo orden de columnas.'
        $null=Invoke-InventoryQueueWorker -QueueDirectory $directory -Destination (New-TestDestination) -Sender $sender -DestinationResolver $resolver
        Assert-Equal $script:ResolveCount 1 'Relee destino sin registros pendientes.'
        Assert-Equal $script:SendCount 1 'Duplico envio tras actualizar destino.'
    }
    Invoke-Test 'Fallo de validacion del destino ocurre antes del POST' {
        $directory=New-TestDirectory; $envelope=New-TestEnvelope
        $null=Add-InventoryQueueRequest $envelope $directory
        $script:SendCount=0
        $receipt=Invoke-InventoryQueueWorker -QueueDirectory $directory -Destination (New-TestDestination) -DestinationResolver { throw 'Columnas modificadas: destino invalido.' } -Sender { $script:SendCount++; [pscustomobject]@{index=0} } -WarningAction SilentlyContinue
        Assert-Equal $receipt.status 'needs-review' 'Oculto columnas invalidas.'
        Assert-Equal $script:SendCount 0 'Envio sin validar el destino.'
        $item=Read-InventoryQueueItem (Get-InventoryQueuePath $directory $envelope.requestId)
        Assert-Equal $item.attemptedAtUtc $null 'Marco intento antes de validar destino.'
        Assert-True ($item.lastError.StartsWith('No se intento')) 'Presento fallo previo como envio ambiguo.'
    }
    Invoke-Test 'API rechaza JSON raiz null, primitivo o array de envelope valido' {
        $directory=New-TestDirectory
        $valid=New-TestEnvelope|ConvertTo-Json -Depth 8 -Compress
        foreach ($json in @('null','true','1',('['+$valid+']'))) {
            $result=Invoke-TestApi $null $directory -Bytes ([Text.Encoding]::UTF8.GetBytes($json))
            Assert-Equal $result.StatusCode 400 ('Acepto JSON sin envelope objeto: '+$json)
        }
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.json').Count 0 'Persistio un envelope con raiz no valida.'
    }
} finally {
    $resolved=[IO.Path]::GetFullPath($script:TempRoot)
    $temp=[IO.Path]::GetFullPath($TestRoot).TrimEnd('\')+'\'
    if ($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved) -match '^it-r-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
Write-Host ("Pruebas receptor correctas: {0}; fallidas: {1}." -f $script:Passed,$script:Failures.Count)
if ($script:Failures.Count) { throw ($script:Failures -join [Environment]::NewLine) }
