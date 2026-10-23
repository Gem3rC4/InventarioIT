#requires -Version 5.1
<#
.SYNOPSIS
Prueba el cliente sin Microsoft, red ni consultas de hardware reales.
#>
[CmdletBinding()]
param([string]$ScriptPath)
$ErrorActionPreference = 'Stop'
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) 'Inventario-Equipo.ps1' }
. $ScriptPath
$script:Passed = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:HttpCalls = New-Object 'System.Collections.Generic.List[object]'
$script:TempRoot = Join-Path ([IO.Path]::GetTempPath()) ('InventarioIT-cliente-tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $script:TempRoot
$script:Key = 'clave-sintetica-pruebas-sin-uso-real-123456'
$script:Handler = { throw 'No hay respuesta HTTP simulada.' }
$script:CollectionCalls = 0
function Assert-True { param([bool]$Condition,[string]$Message); if (-not $Condition) { throw $Message } }
function Assert-Equal {
    param([AllowNull()][object]$Actual,[AllowNull()][object]$Expected,[string]$Message)
    if ($null -eq $Expected) { if ($null -ne $Actual) { throw "$Message (se esperaba null)" } }
    elseif ($Actual -cne $Expected) { throw "$Message (esperado: $Expected; recibido: $Actual)" }
}
function Assert-Throws {
    param([scriptblock]$Action,[string]$Message)
    $thrown = $false
    try { & $Action | Out-Null } catch { $thrown = $true }
    Assert-True $thrown $Message
}
function Invoke-Test {
    param([string]$Name,[scriptblock]$Action)
    try { & $Action; $script:Passed++; Write-Host "OK: $Name" -ForegroundColor Green }
    catch { $script:Failures.Add(("{0}: {1}" -f $Name,$_.Exception.Message)); Write-Host ("FALLO: {0}: {1}" -f $Name,$_.Exception.Message) -ForegroundColor Red }
}
function New-TestRecord {
    $record = [ordered]@{}
    foreach ($field in Get-InventoryFields) { $record[$field] = 'valor-' + $field }
    $record.Equipo = 'PC-PRUEBA'
    $record.Area = 'Tecnolog' + [char]0x00ED + 'a'
    $record.RAM_GB = 16.25; $record.DiscoTotal_GB = 256; $record.DiscoUsado_GB = 64.5; $record.DiscoLibre_GB = $null
    [pscustomobject]$record
}
function Get-InventoryRecord { param([string]$Area); $script:CollectionCalls++; New-TestRecord }
function Invoke-RestMethod {
    [CmdletBinding()]
    param([string]$Method,[string]$Uri,[hashtable]$Headers,[string]$ContentType,[object]$Body,[int]$TimeoutSec,[int]$MaximumRedirection)
    $script:HttpCalls.Add([pscustomobject]@{ Method=$Method; Uri=$Uri; Headers=$Headers; Body=$Body; ContentType=$ContentType; Redirects=$MaximumRedirection })
    & $script:Handler $Method $Uri $Body
}
function Start-Sleep { param([int]$Milliseconds) }
function Set-ResponseMock {
    param([string]$PostStatus='queued',[string]$GetStatus='sent',[switch]$FailPost,[switch]$FailGet,[switch]$WrongId)
    $script:PostStatus=$PostStatus; $script:GetStatus=$GetStatus; $script:FailPost=$FailPost.IsPresent
    $script:FailGet=$FailGet.IsPresent; $script:WrongId=$WrongId.IsPresent; $script:HttpCalls.Clear()
    $script:Handler = {
        param($Method,$Uri,$Body)
        if ($Method -eq 'Post') {
            if ($script:FailPost) { throw 'Fallo HTTP sintetico.' }
            Assert-True ($Body -is [byte[]]) 'El POST debe transmitir bytes UTF-8.'
            $payload = [Text.Encoding]::UTF8.GetString($Body) | ConvertFrom-Json
            $script:LastEnvelope=$payload
            $id=$payload.requestId
            if ($script:WrongId) { $id=[guid]::NewGuid().ToString('D') }
            return [pscustomobject]@{ status=$script:PostStatus; requestId=$id }
        }
        if ($script:FailGet) { throw 'Estado temporalmente inaccesible.' }
        [pscustomobject]@{ status=$script:GetStatus; requestId=$script:LastEnvelope.requestId }
    }
}
try {
    Invoke-Test 'Envelope contiene version, ID estable y los18 campos' {
        $record=New-TestRecord
        $a=New-InventoryEnvelope $record; $b=New-InventoryEnvelope $record
        Assert-InventoryEnvelope $a
        Assert-Equal @($a.PSObject.Properties.Name).Count 3 'Cantidad de propiedades incorrecta.'
        Assert-True ($a.requestId -ne $b.requestId) 'Capturas distintas reutilizan ID.'
        $guid=[guid]::Empty
        Assert-True ([guid]::TryParseExact($a.requestId,'D',[ref]$guid)) 'ID no es GUID canonico.'
        Assert-Equal @($a.record.PSObject.Properties.Name).Count 18 'Perdio campos del registro.'
    }
    Invoke-Test 'Rechaza version booleana, texto, ID ajeno y registro incompleto' {
        foreach ($version in @($true,'1',2,$null)) {
            $envelope=New-InventoryEnvelope (New-TestRecord); $envelope.schemaVersion=$version
            Assert-Throws { Assert-InventoryEnvelope $envelope } 'Acepto version invalida.'
        }
        $envelope=New-InventoryEnvelope (New-TestRecord); $envelope.requestId='..\archivo'
        Assert-Throws { Assert-InventoryEnvelope $envelope } 'Acepto ID que no es GUID.'
        $envelope=New-InventoryEnvelope (New-TestRecord); $envelope.record.PSObject.Properties.Remove('Area')
        Assert-Throws { Assert-InventoryEnvelope $envelope } 'Acepto registro incompleto.'
    }
    Invoke-Test 'URL privada permitida y credenciales/rutas/HTTP publico bloqueados' {
        foreach ($url in @('http://127.0.0.1:8765','http://192.168.1.5:8765','http://10.1.1.5:8765','https://receptor.example.invalid')) { Assert-ReceiverUrl $url }
        foreach ($url in @('http://8.8.8.8','http://public.example.invalid','https://user:password@receptor.example.invalid','https://receptor.example.invalid/otra','https://receptor.example.invalid/?key=x','https://receptor.example.invalid/#x')) {
            Assert-Throws { Assert-ReceiverUrl $url } 'Acepto URL insegura o con componentes indebidos.'
        }
    }
    Invoke-Test 'POST usa bytes UTF-8, clave propia, redireccion cero y conserva datos' {
        Set-ResponseMock
        $envelope=New-InventoryEnvelope (New-TestRecord)
        $response=Send-InventoryEnvelope $envelope 'https://receptor.example.invalid' $script:Key 0
        Assert-Equal $response.status 'queued' 'Cola se interpreto como envio confirmado.'
        Assert-Equal $script:HttpCalls.Count 1 'Ejecuto mas llamadas de las solicitadas.'
        $call=$script:HttpCalls[0]
        Assert-Equal $call.Uri 'https://receptor.example.invalid/inventario' 'Ruta incorrecta.'
        Assert-Equal $call.Headers.Authorization ('Bearer '+$script:Key) 'Cabecera de acceso incorrecta.'
        Assert-Equal $call.Redirects 0 'Puede reenviar la clave a otro host.'
        Assert-Equal $script:LastEnvelope.record.Area $envelope.record.Area 'Perdio Unicode.'
        Assert-Equal $script:LastEnvelope.record.RAM_GB 16.25 'Perdio numero decimal.'
        Assert-Equal $script:LastEnvelope.record.DiscoLibre_GB $null 'Perdio null.'
    }
    Invoke-Test 'Consulta de estado confirma sent con un solo POST' {
        Set-ResponseMock
        $response=Send-InventoryEnvelope (New-InventoryEnvelope (New-TestRecord)) 'https://receptor.example.invalid' $script:Key 15
        Assert-Equal $response.status 'sent' 'No obtuvo confirmacion.'
        Assert-Equal @($script:HttpCalls|Where-Object{$_.Method -eq 'Post'}).Count 1 'Repitio POST mientras consultaba estado.'
        Assert-Equal @($script:HttpCalls|Where-Object{$_.Method -eq 'Get'}).Count 1 'Consulta de estado incorrecta.'
    }
    Invoke-Test 'Fallo POST no se reintenta y respuesta de otro ID se rechaza' {
        Set-ResponseMock -FailPost
        Assert-Throws { Send-InventoryEnvelope (New-InventoryEnvelope (New-TestRecord)) 'https://receptor.example.invalid' $script:Key 0 } 'Oculto fallo POST.'
        Assert-Equal $script:HttpCalls.Count 1 'Reintento POST automaticamente.'
        Set-ResponseMock -WrongId
        Assert-Throws { Send-InventoryEnvelope (New-InventoryEnvelope (New-TestRecord)) 'https://receptor.example.invalid' $script:Key 0 } 'Acepto recibo de otra solicitud.'
    }
    Invoke-Test 'Fallo GET conserva estado queued sin volver a publicar' {
        Set-ResponseMock -FailGet
        $response=Send-InventoryEnvelope (New-InventoryEnvelope (New-TestRecord)) 'https://receptor.example.invalid' $script:Key 15
        Assert-Equal $response.status 'queued' 'Un fallo de consulta cambio el estado aceptado.'
        Assert-Equal @($script:HttpCalls|Where-Object{$_.Method -eq 'Post'}).Count 1 'Repitio POST por fallo de consulta.'
    }
    Invoke-Test 'SoloLocal persiste solicitud recuperable y no usa red' {
        Set-ResponseMock
        $result=Invoke-InventoryMain -Area TI -DirectorioSalida (Join-Path $script:TempRoot 'local') -SoloLocal
        Assert-Equal $result.Estado 'SoloLocal' 'Estado local incorrecto.'
        Assert-Equal $script:HttpCalls.Count 0 'SoloLocal uso red.'
        $saved=Get-Content -LiteralPath $result.Solicitud -Raw -Encoding UTF8|ConvertFrom-Json
        Assert-InventoryEnvelope $saved
        Assert-Equal $saved.record.Area (New-TestRecord).Area 'Respaldo cambio Unicode.'
    }
    Invoke-Test 'Reenvio SolicitudJson reutiliza ID sin recopilar otra vez' {
        $envelope=New-InventoryEnvelope (New-TestRecord)
        $path=Join-Path $script:TempRoot 'original.solicitud.json'
        $envelope|ConvertTo-Json -Depth 6|Set-Content -LiteralPath $path -Encoding UTF8
        $before=$script:CollectionCalls
        Set-ResponseMock -PostStatus sent
        $result=Invoke-InventoryMain -SolicitudJson $path -ServidorUrl 'https://receptor.example.invalid' -ClaveApi $script:Key -DirectorioSalida (Join-Path $script:TempRoot 'reenvio') -EsperarSegundos 0
        Assert-Equal $result.Estado 'RegistradoEnExcel' 'No distinguio Excel confirmado.'
        Assert-Equal $result.SolicitudId $envelope.requestId 'Creo otro ID al reenviar.'
        Assert-Equal $script:CollectionCalls $before 'Recopilo otra captura al reenviar.'
        Assert-Equal $script:LastEnvelope.requestId $envelope.requestId 'Transmitio otro ID.'
    }
    Invoke-Test 'Main queued queda EnCola y guarda solicitud antes de fallo POST' {
        Set-ResponseMock
        $result=Invoke-InventoryMain -Area TI -ServidorUrl 'https://receptor.example.invalid' -ClaveApi $script:Key -DirectorioSalida (Join-Path $script:TempRoot 'cola') -EsperarSegundos 0
        Assert-Equal $result.Estado 'EnCola' 'Presento queued como registrado en Excel.'
        Set-ResponseMock -FailPost
        $directory=Join-Path $script:TempRoot 'fallo'
        $oldTls=[Net.ServicePointManager]::SecurityProtocol
        Assert-Throws { Invoke-InventoryMain -Area TI -ServidorUrl 'https://receptor.example.invalid' -ClaveApi $script:Key -DirectorioSalida $directory -EsperarSegundos 0 -WarningAction SilentlyContinue } 'Oculto fallo POST.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.solicitud.json').Count 1 'No conservo solicitud al fallar.'
        Assert-Equal ([Net.ServicePointManager]::SecurityProtocol) $oldTls 'No restauro TLS.'
    }
    Invoke-Test 'Main needs-review informa error y conserva mismo ID recuperable' {
        Set-ResponseMock -PostStatus needs-review
        $directory=Join-Path $script:TempRoot 'revision'
        Assert-Throws { Invoke-InventoryMain -Area TI -ServidorUrl 'https://receptor.example.invalid' -ClaveApi $script:Key -DirectorioSalida $directory -EsperarSegundos 0 -WarningAction SilentlyContinue } 'Presento revision como exito.'
        Assert-Equal @($script:HttpCalls|Where-Object{$_.Method -eq 'Post'}).Count 1 'Reenvio un registro en revision.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.solicitud.json').Count 1 'Perdio solicitud de revision.'
    }
} finally {
    $resolved=[IO.Path]::GetFullPath($script:TempRoot)
    $temp=[IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\')+'\'
    if ($resolved.StartsWith($temp,[StringComparison]::OrdinalIgnoreCase) -and [IO.Path]::GetFileName($resolved) -match '^InventarioIT-cliente-tests-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
Write-Host ("Pruebas cliente correctas: {0}; fallidas: {1}." -f $script:Passed,$script:Failures.Count)
if ($script:Failures.Count) { throw ($script:Failures -join [Environment]::NewLine) }
