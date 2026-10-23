#requires -Version 5.1
<#
.SYNOPSIS
Prueba el inicio de sesion de Microsoft del receptor con respuestas simuladas.
No se conecta a Microsoft ni a Excel.
#>
[CmdletBinding()]
param([string]$ModulePath)
$ErrorActionPreference = 'Stop'
if (-not $ModulePath) { $ModulePath = Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) 'Inventario.Funciones.psm1' }
$module = Import-Module $ModulePath -Force -PassThru
$script:Passed = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:TempRoot = Join-Path ([IO.Path]::GetTempPath()) ('InventarioIT-sesion-tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $script:TempRoot
function Assert-True { param([bool]$Condition,[string]$Message); if (-not $Condition) { throw $Message } }
function Assert-Equal {
    param([AllowNull()][object]$Actual,[AllowNull()][object]$Expected,[string]$Message)
    if ($Actual -cne $Expected) { throw "$Message (esperado: $Expected; recibido: $Actual)" }
}
function Invoke-Test {
    param([string]$Name,[scriptblock]$Action)
    try { & $Action; $script:Passed++; Write-Host "OK: $Name" -ForegroundColor Green }
    catch { $script:Failures.Add(("{0}: {1}" -f $Name,$_.Exception.Message)); Write-Host ("FALLO: {0}: {1}" -f $Name,$_.Exception.Message) -ForegroundColor Red }
}

# Reemplaza Invoke-RestMethod y utilidades de escritorio DENTRO del modulo.
& $module {
    $script:Calls = New-Object 'System.Collections.Generic.List[object]'
    $script:Responses = New-Object 'System.Collections.Generic.Queue[object]'
    function script:Invoke-RestMethod {
        [CmdletBinding()]
        param([string]$Method,[string]$Uri,[object]$Body,[string]$ContentType,[hashtable]$Headers,[int]$TimeoutSec,[switch]$UseBasicParsing)
        $script:Calls.Add([pscustomobject]@{ Method=$Method; Uri=$Uri; Body=$Body; ContentType=$ContentType; Headers=$Headers })
        if ($script:Responses.Count -eq 0) { throw 'No hay respuesta simulada.' }
        $next = $script:Responses.Dequeue()
        if ($next -is [hashtable] -and $next.ContainsKey('ErrorJson')) {
            $record = New-Object Management.Automation.ErrorRecord((New-Object Exception 'HTTP simulado'),'HttpSimulado','InvalidOperation',$null)
            $record.ErrorDetails = New-Object Management.Automation.ErrorDetails($next.ErrorJson)
            $PSCmdlet.ThrowTerminatingError($record)
        }
        return $next
    }
    function script:Start-Sleep { param([int]$Seconds,[int]$Milliseconds) }
    function script:Start-Process { param([string]$FilePath) }
    function script:Set-Clipboard { param([string]$Value) }
    function script:Write-Host { param([object]$Object,[object]$ForegroundColor) }
}
function Reset-Mock { param([object[]]$Responses) & $module { param($r) $script:Calls.Clear(); $script:Responses.Clear(); foreach ($x in $r) { $script:Responses.Enqueue($x) } } $Responses }
function Get-Calls { & $module { ,$script:Calls.ToArray() } }
function Get-Session { & $module { $script:GraphSession } }
$deviceResponse = [pscustomobject]@{ user_code='ABCD-1234'; device_code='dev-123'; verification_uri='https://microsoft.com/devicelogin'; interval=5; expires_in=900 }
$tokenResponse = [pscustomobject]@{ access_token='acceso-1'; refresh_token='renovacion-secreta-1'; expires_in=3600 }

try {
    Invoke-Test 'Codigo de dispositivo espera authorization_pending y guarda la sesion cifrada' {
        $cache = Join-Path $script:TempRoot 'sesion1.dat'
        Reset-Mock @($deviceResponse, @{ ErrorJson='{"error":"authorization_pending","error_description":"AADSTS70016: pendiente"}' }, $tokenResponse)
        Connect-InventarioGraph -CachePath $cache
        $calls = Get-Calls
        Assert-Equal $calls.Count 3 'Cantidad de llamadas.'
        Assert-True ($calls[0].Uri -like 'https://login.microsoftonline.com/organizations/oauth2/v2.0/devicecode') 'No pidio el codigo al tenant organizations.'
        Assert-True ($calls[0].Body.scope -match 'Files\.ReadWrite' -and $calls[0].Body.scope -match 'offline_access') 'Permisos incorrectos.'
        Assert-Equal $calls[2].Body.grant_type 'urn:ietf:params:oauth:grant-type:device_code' 'Tipo de concesion.'
        Assert-Equal (Get-Session).AccessToken 'acceso-1' 'No guardo el token en memoria.'
        $content = [IO.File]::ReadAllText($cache)
        Assert-True ($content -notmatch 'renovacion-secreta-1') 'El archivo de sesion guardo el token sin cifrar.'
    }
    Invoke-Test 'Reinicio reutiliza la sesion guardada sin pedir codigo' {
        $cache = Join-Path $script:TempRoot 'sesion1.dat'
        Reset-Mock @([pscustomobject]@{ access_token='acceso-2'; refresh_token='renovacion-secreta-2'; expires_in=3600 })
        Connect-InventarioGraph -CachePath $cache
        $calls = Get-Calls
        Assert-Equal $calls.Count 1 'Pidio un codigo nuevo.'
        Assert-Equal $calls[0].Body.grant_type 'refresh_token' 'No uso la sesion guardada.'
        Assert-Equal $calls[0].Body.refresh_token 'renovacion-secreta-1' 'No descifro la sesion guardada.'
        Assert-Equal (Get-Session).RefreshToken 'renovacion-secreta-2' 'No roto el token de renovacion.'
    }
    Invoke-Test 'NuevaSesion ignora la sesion guardada' {
        $cache = Join-Path $script:TempRoot 'sesion1.dat'
        Reset-Mock @($deviceResponse, $tokenResponse)
        Connect-InventarioGraph -CachePath $cache -NuevaSesion
        Assert-True ((Get-Calls)[0].Uri -like '*devicecode') 'No pidio codigo nuevo.'
    }
    Invoke-Test 'Sesion guardada invalida vuelve al codigo de dispositivo' {
        $cache = Join-Path $script:TempRoot 'sesion1.dat'
        Reset-Mock @(@{ ErrorJson='{"error":"invalid_grant","error_description":"AADSTS700082: expirado"}' }, $deviceResponse, $tokenResponse)
        Connect-InventarioGraph -CachePath $cache
        Assert-Equal (Get-Calls).Count 3 'No recupero el inicio de sesion.'
    }
    Invoke-Test 'Rechazo en el navegador detiene con mensaje claro' {
        Reset-Mock @($deviceResponse, @{ ErrorJson='{"error":"authorization_declined"}' })
        $message = $null
        try { Connect-InventarioGraph -CachePath (Join-Path $script:TempRoot 'sesion2.dat') } catch { $message = $_.Exception.Message }
        Assert-True ($message -match 'cancelado') 'No informo la cancelacion.'
    }
    Invoke-Test 'Consulta Graph envia Bearer y cuerpo UTF-8; renueva si expiro' {
        Reset-Mock @($deviceResponse, $tokenResponse)
        Connect-InventarioGraph -CachePath (Join-Path $script:TempRoot 'sesion3.dat')
        & $module { $script:GraphSession.ExpiresAtUtc = [DateTime]::UtcNow.AddMinutes(-1) }
        Reset-Mock @([pscustomobject]@{ access_token='acceso-nuevo'; expires_in=3600 }, [pscustomobject]@{ index=7 })
        $body = '{"values":[["Tecnolog' + [char]0x00ED + 'a"]]}'
        $result = Invoke-InventarioGraphRequest -Method POST -Uri 'https://graph.microsoft.com/v1.0/drives/D/items/I/workbook/tables/T/rows/add' -Body $body
        $calls = Get-Calls
        Assert-Equal $result.index 7 'No devolvio la respuesta de Graph.'
        Assert-Equal $calls[0].Body.grant_type 'refresh_token' 'No renovo el token vencido.'
        Assert-Equal $calls[1].Headers.Authorization 'Bearer acceso-nuevo' 'Uso un token viejo.'
        Assert-Equal ([Text.Encoding]::UTF8.GetString($calls[1].Body)) $body 'El cuerpo no viajo en UTF-8.'
        Assert-Equal (Get-Session).RefreshToken 'renovacion-secreta-1' 'Perdio el token de renovacion al no recibir uno nuevo.'
    }
    Invoke-Test 'Error de Graph se informa con codigo y mensaje' {
        Reset-Mock @(@{ ErrorJson='{"error":{"code":"accessDenied","message":"Acceso denegado"}}' })
        $message = $null
        try { $null = Invoke-InventarioGraphRequest -Uri 'https://graph.microsoft.com/v1.0/me' } catch { $message = $_.Exception.Message }
        Assert-True ($message -match 'accessDenied' -and $message -match 'Acceso denegado') "Mensaje poco claro: $message"
    }
    Invoke-Test 'No envia el token a otros hosts' {
        Reset-Mock @()
        $thrown = $false
        try { $null = Invoke-InventarioGraphRequest -Uri 'https://ejemplo.invalid/v1.0/me' } catch { $thrown = $true }
        Assert-True $thrown 'Acepto un host distinto de Graph.'
        Assert-Equal (Get-Calls).Count 0 'Hizo una llamada externa.'
    }
    Invoke-Test 'Reconoce Excel personal, direccion del navegador y SharePoint' {
        $direct = Get-ExcelUrlInfo 'https://excel.cloud.microsoft/open/onedrive/?docId=ABCDEF0123456789%21s0123456789abcdef0123456789abcdef&driveId=ABCDEF0123456789'
        Assert-Equal $direct.Kind 'Direct' 'No uso los identificadores del navegador.'
        Assert-Equal $direct.ItemId 'ABCDEF0123456789!s0123456789abcdef0123456789abcdef' 'No decodifico docId.'
        Assert-Equal $direct.DriveId 'ABCDEF0123456789' 'driveId.'
        Assert-Equal $direct.Personal $true 'No detecto OneDrive personal.'
        $share = Get-ExcelUrlInfo 'https://1drv.ms/x/c/abcdef0123456789/IQAAAAAAAAAAAAAAAAAAAAAAAAA?e=abc123'
        Assert-Equal "$($share.Kind)/$($share.Personal)" 'Share/True' '1drv.ms debe ser enlace personal.'
        $business = Get-ExcelUrlInfo 'https://contoso-my.sharepoint.com/:x:/g/personal/usuario_contoso_com/EjemploToken?e=abc'
        Assert-Equal "$($business.Kind)/$($business.Personal)" 'Share/False' 'SharePoint no es cuenta personal.'
        $thrown = $false
        try { $null = Get-ExcelUrlInfo 'https://ejemplo.invalid/libro.xlsx' } catch { $thrown = $true }
        Assert-True $thrown 'Acepto un sitio que no es de Microsoft.'
    }
    Invoke-Test 'Destino directo consulta las tablas del archivo sin usar shares' {
        & $module {
            $script:GraphUris = New-Object 'System.Collections.Generic.List[string]'
            function script:Invoke-InventarioGraphRequest {
                param([string]$Method='GET',[string]$Uri,[string]$Body)
                $script:GraphUris.Add($Uri)
                if ($Uri -match '/tables$') { return [pscustomobject]@{ value = @([pscustomobject]@{ id='{T1}'; name='Tabla1' }) } }
                $i = 0
                return [pscustomobject]@{ value = @(Get-InventoryFields | ForEach-Object { [pscustomobject]@{ name=$_; index=$i }; $i++ }) }
            }
        }
        $destination = Get-ExcelDestination -Url 'https://excel.cloud.microsoft/open/onedrive/?docId=ABCDEF0123456789%21s0123&driveId=ABCDEF0123456789'
        $uris = & $module { ,$script:GraphUris.ToArray() }
        Assert-True ($uris[0] -in @('https://graph.microsoft.com/v1.0/drives/ABCDEF0123456789/items/ABCDEF0123456789!s0123/workbook/tables','https://graph.microsoft.com/v1.0/drives/ABCDEF0123456789/items/ABCDEF0123456789%21s0123/workbook/tables')) "Ruta de tablas incorrecta: $($uris[0])"
        Assert-True (@($uris | Where-Object { $_ -match '/shares/' }).Count -eq 0) 'Uso shares con identificadores directos.'
        Assert-Equal $destination.TableName 'Tabla1' 'No selecciono la unica tabla.'
        Assert-Equal @($destination.Columns).Count 18 'Columnas.'
    }
} finally {
    $resolved = [IO.Path]::GetFullPath($script:TempRoot)
    if ([IO.Path]::GetFileName($resolved) -match '^InventarioIT-sesion-tests-[a-f0-9]{32}$') { Remove-Item -LiteralPath $resolved -Recurse -Force }
}
Write-Host ("Pruebas sesion correctas: {0}; fallidas: {1}." -f $script:Passed,$script:Failures.Count)
if ($script:Failures.Count) { throw ($script:Failures -join [Environment]::NewLine) }
