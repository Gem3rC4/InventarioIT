#requires -Version 5.1
<#
.SYNOPSIS
Pruebas sin dependencias, conexiones ni cambios en Excel, para PowerShell 5.1.
#>
[CmdletBinding()]
param([string]$ScriptPath)
if (-not $ScriptPath) { $ScriptPath = Join-Path (Split-Path (Split-Path $MyInvocation.MyCommand.Path -Parent) -Parent) 'Inventario-Equipo.ps1' }
$ErrorActionPreference = 'Stop'
. $ScriptPath
$script:Passed = 0
$script:Failures = New-Object 'System.Collections.Generic.List[string]'
$script:Calls = New-Object 'System.Collections.Generic.List[object]'
$script:Contexts = New-Object 'System.Collections.Generic.List[object]'
$script:Handler = { throw 'Graph no tiene respuesta simulada.' }
$script:Connections = 0
$script:Disconnections = 0
$script:TempRoot = Join-Path ([IO.Path]::GetTempPath()) ('InventarioIT-tests-' + [guid]::NewGuid().ToString('N'))
$null = New-Item -ItemType Directory -Path $script:TempRoot
function Assert-True {
    param([bool]$Condition, [string]$Message)
    if (-not $Condition) { throw $Message }
}
function Assert-Equal {
    param([AllowNull()][object]$Actual, [AllowNull()][object]$Expected, [string]$Message)
    if ($null -eq $Expected) {
        if ($null -ne $Actual) { throw "$Message (se esperaba null)" }
    } elseif ($Actual -cne $Expected) { throw "$Message (esperado: $Expected; recibido: $Actual)" }
}
function Assert-Throws {
    param([scriptblock]$Action, [string]$Message)
    $didThrow = $false
    try { & $Action | Out-Null } catch { $didThrow = $true }
    Assert-True $didThrow $Message
}
function Invoke-Test {
    param([string]$Name, [scriptblock]$Action)
    try { & $Action; $script:Passed++; Write-Host "OK: $Name" -ForegroundColor Green }
    catch {
        $script:Failures.Add(("{0}: {1}" -f $Name, $_.Exception.Message))
        Write-Host ("FALLO: {0}: {1}" -f $Name, $_.Exception.Message) -ForegroundColor Red
    }
}
function New-TestRecord {
    $record = [ordered]@{}
    foreach ($name in Get-InventoryFields) { $record[$name] = 'valor-' + $name }
    $record.Equipo = 'PC-PRUEBA'
    $record.Area = 'Tecnolog' + [char]0x00ED + 'a'
    $record.RAM_GB = 16.25
    $record.DiscoTotal_GB = 512
    $record.DiscoUsado_GB = 128.5
    $record.DiscoLibre_GB = $null
    [pscustomobject]$record
}
function New-TestColumns {
    param([string[]]$Names = (Get-InventoryFields))
    for ($i = 0; $i -lt $Names.Count; $i++) { [pscustomobject]@{ index = $i; name = $Names[$i] } }
}
# Se reemplazan todas las operaciones de Microsoft Graph por respuestas sinteticas.
function Invoke-MgGraphRequest {
    [CmdletBinding()]
    param([string]$Method, [string]$Uri, [object]$Body, [string]$ContentType, [string]$OutputType)
    $script:Calls.Add([pscustomobject]@{ Method = $Method; Uri = $Uri; Body = $Body; ContentType = $ContentType })
    & $script:Handler $Method $Uri $Body
}
function Initialize-GraphModule { param([switch]$Install) }
function Get-MgContext { return $null }
function Connect-MgGraph {
    [CmdletBinding()]
    param([string[]]$Scopes, [string]$ContextScope, [switch]$UseDeviceAuthentication, [switch]$NoWelcome,
        [string]$TenantId, [string]$ClientId)
    Assert-Equal ($Scopes -join ',') 'Files.ReadWrite' 'Permiso delegado incorrecto.'
    Assert-Equal $ContextScope 'Process' 'La sesion no debe persistir.'
    $script:Connections++
}
$script:MutableContext = [pscustomobject]@{ MaxRetry = 3; ClientTimeout = [timespan]::FromSeconds(300) }
function Get-MgRequestContext { $script:MutableContext }
function Set-MgRequestContext {
    param([int]$MaxRetry, [int]$ClientTimeout)
    $script:Contexts.Add([pscustomobject]@{ MaxRetry = $MaxRetry; ClientTimeout = $ClientTimeout })
    $script:MutableContext.MaxRetry = $MaxRetry
    $script:MutableContext.ClientTimeout = [timespan]::FromSeconds($ClientTimeout)
}
function Disconnect-MgGraph { [CmdletBinding()] param(); $script:Disconnections++ }
$script:SharedUrl = 'https://example-my.sharepoint.com/:x:/g/personal/test/ID?e=abc'
function Set-DestinationMock {
    param([object[]]$Columns, [switch]$Remote, [switch]$FailPost, [switch]$UnconfirmedPost, [int]$TableCount = 1)
    $script:Columns = $Columns
    $script:Remote = $Remote.IsPresent
    $script:FailPost = $FailPost.IsPresent
    $script:UnconfirmedPost = $UnconfirmedPost.IsPresent
    $script:TableCount = $TableCount
    $script:Calls.Clear()
    $script:Handler = {
        param($Method, $Uri, $Body)
        if ($Method -eq 'POST') {
            Assert-True ($script:Contexts.Count -gt 0) 'No configuro contexto antes del POST.'
            Assert-Equal $script:Contexts[$script:Contexts.Count - 1].MaxRetry 0 'El POST permitio reintentos del SDK.'
            if ($script:FailPost) { throw 'HTTP 504 simulado: resultado desconocido.' }
            if ($script:UnconfirmedPost) { return [pscustomobject]@{ values = @() } }
            return [pscustomobject]@{ index = 0 }
        }
        if ($Uri -match '/shares/.+/driveItem$') {
            $item = [pscustomobject]@{ id = 'file/id'; parentReference = [pscustomobject]@{ driveId = 'drive+id' } }
            if ($script:Remote) { return [pscustomobject]@{ id = 'wrapper'; remoteItem = $item } }
            return $item
        }
        if ($Uri -match '/workbook/tables$') {
            $tables = @()
            for ($i = 0; $i -lt $script:TableCount; $i++) { $tables += [pscustomobject]@{ id = 'table{' + $i + '}'; name = 'Inventario' + $i } }
            return [pscustomobject]@{ value = $tables }
        }
        if ($Uri -match '/columns$') { return [pscustomobject]@{ value = $script:Columns } }
        throw ('Solicitud Graph inesperada: ' + $Uri)
    }
}
try {
    Invoke-Test 'Esquema exacto de 18 campos' {
        $expected = 'Equipo,Usuario,Marca,Modelo,Serie,Procesador,TipoEquipo,RAM_GB,DiscoTotal_GB,DiscoUsado_GB,DiscoLibre_GB,TipoDisco,IP,Windows,Version,Activado,Fecha,Area'
        Assert-Equal ((Get-InventoryFields) -join ',') $expected 'Esquema incorrecto.'
        Assert-InventoryRecord (New-TestRecord)
    }
    Invoke-Test 'Rechaza registros incompletos, adicionales y valores invalidos' {
        $missing = New-TestRecord
        $missing.PSObject.Properties.Remove('Area')
        Assert-Throws { Assert-InventoryRecord $missing } 'Acepto campo faltante.'
        $extra = New-TestRecord
        $extra | Add-Member Nota 'extra'
        Assert-Throws { Assert-InventoryRecord $extra } 'Acepto campo adicional.'
        $nested = New-TestRecord
        $nested.Area = @('uno','dos')
        Assert-Throws { Assert-InventoryRecord $nested } 'Acepto valor no escalar.'
        foreach ($value in @(-1, '16', $true, [double]::NaN, [double]::PositiveInfinity)) {
            $bad = New-TestRecord
            $bad.RAM_GB = $value
            Assert-Throws { Assert-InventoryRecord $bad } 'Acepto numero invalido.'
        }
    }
    Invoke-Test 'Neutraliza formulas y conserva numeros, null y Unicode' {
        foreach ($value in @('=1+1','+cmd','-cmd','@SUM(A1)', '  =1', ([string][char]9 + '=1'))) {
            Assert-Equal (ConvertTo-ExcelSafeValue $value) ("'" + $value) 'Formula sin neutralizar.'
        }
        Assert-Equal (ConvertTo-ExcelSafeValue 'normal') 'normal' 'Modifico texto seguro.'
        Assert-Equal (ConvertTo-ExcelSafeValue -Value ([double]-1.5)) -1.5 'Modifico numero negativo.'
        Assert-Equal (ConvertTo-ExcelSafeValue $null) $null 'Modifico null.'
        Assert-Equal (ConvertTo-ExcelSafeValue (New-TestRecord).Area) (New-TestRecord).Area 'Modifico Unicode.'
    }
    Invoke-Test 'JSON 1 x 18 conserva orden de columnas, null y numeros' {
        $record = New-TestRecord
        $record.Usuario = '=1+1'
        $names = @(Get-InventoryFields)
        [array]::Reverse($names)
        $json = New-InventoryRowBody $record @(New-TestColumns $names)
        Assert-True ($json -match '"values":\[\[') ('La matriz no es 1 x 18: ' + $json)
        $body = $json | ConvertFrom-Json
        Assert-Equal $body.values.Count 1 'Cantidad de filas incorrecta.'
        Assert-Equal $body.values[0].Count 18 'Null u otro dato desaparecio.'
        for ($i = 0; $i -lt $names.Count; $i++) {
            $expected = ConvertTo-ExcelSafeValue $record.PSObject.Properties[$names[$i]].Value
            Assert-Equal $body.values[0][$i] $expected ('Valor fuera de columna: ' + $names[$i])
        }
        Assert-True ($body.values[0][9] -is [ValueType]) 'Los GB dejaron de ser numeros.'
    }
    Invoke-Test 'Respaldo JSON original y CSV protegido' {
        $record = New-TestRecord
        $record.Equipo = '=EQUIPO'
        $record.Area = '@AREA'
        $backup = Save-InventoryBackup $record (Join-Path $script:TempRoot 'respaldos')
        $restored = Get-Content -LiteralPath $backup.Json -Raw -Encoding UTF8 | ConvertFrom-Json
        Assert-Equal $restored.Equipo '=EQUIPO' 'Modifico JSON original.'
        Assert-Equal $restored.DiscoLibre_GB $null 'Perdio null en JSON.'
        $csv = Import-Csv -LiteralPath $backup.Csv -Encoding UTF8
        Assert-Equal $csv.Equipo "'=EQUIPO" 'CSV permite formula.'
        Assert-Equal $csv.Area "'@AREA" 'CSV permite formula en Area.'
    }
    Invoke-Test 'Token de enlace base64url UTF-8 y validacion de dominio' {
        $url = 'https://example.sharepoint.com/:x:/g/carpeta/' + [char]0x00F1 + '?e=AB+C'
        $token = ConvertTo-SharingToken $url
        Assert-True ($token.StartsWith('u!')) 'Falta prefijo u!.'
        Assert-True ($token -notmatch '[=/+]') 'Token no es base64url sin relleno.'
        $encoded = $token.Substring(2).Replace('-','+').Replace('_','/')
        while ($encoded.Length % 4) { $encoded += '=' }
        Assert-Equal ([Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($encoded))) $url 'UTF-8 incorrecto.'
        foreach ($badUrl in @('http://example.sharepoint.com/x','https://sharepoint.com.evil.invalid/x','https://example.invalid/x')) {
            Assert-Throws { ConvertTo-SharingToken $badUrl } 'Acepto enlace inseguro o ajeno.'
        }
    }
    Invoke-Test 'Resuelve driveItem y remoteItem con IDs escapados' {
        foreach ($remote in @($false, $true)) {
            Set-DestinationMock -Columns @(New-TestColumns) -Remote:$remote
            $destination = Get-ExcelDestination $script:SharedUrl ''
            Assert-True ($destination.TableUri -match '/drives/drive%2Bid/items/file%2Fid/') 'IDs incorrectos o sin escapar.'
            Assert-Equal $destination.TableName 'Inventario0' 'Tabla incorrecta.'
        }
    }
    Invoke-Test 'Acepta columnas reordenadas y conserva correspondencia' {
        $names = @(Get-InventoryFields)
        [array]::Reverse($names)
        Set-DestinationMock -Columns @(New-TestColumns $names)
        $destination = Get-ExcelDestination $script:SharedUrl ''
        Assert-Equal $destination.Columns[0].name 'Area' 'Perdio orden Excel.'
        $body = New-InventoryRowBody (New-TestRecord) $destination.Columns | ConvertFrom-Json
        Assert-Equal $body.values[0][0] (New-TestRecord).Area 'Valor no corresponde a Area.'
    }
    Invoke-Test 'Rechaza encabezados faltantes, adicionales, duplicados o distintos' {
        $normal = @(Get-InventoryFields)
        $missing = @($normal[0..16])
        $extra = @($normal) + 'Nota'
        $duplicate = @($normal); $duplicate[17] = 'Equipo'
        $wrongCase = @($normal); $wrongCase[17] = 'area'
        foreach ($names in @($missing, $extra, $duplicate, $wrongCase)) {
            Set-DestinationMock -Columns @(New-TestColumns $names)
            Assert-Throws { Get-ExcelDestination $script:SharedUrl '' } 'Acepto esquema incorrecto.'
            Assert-Equal @($script:Calls | Where-Object { $_.Method -eq 'POST' }).Count 0 'Inserto con esquema incorrecto.'
        }
        $columns = @(New-TestColumns); $columns[17].index = 16
        Set-DestinationMock -Columns $columns
        Assert-Throws { Get-ExcelDestination $script:SharedUrl '' } 'Acepto indices duplicados.'
    }
    Invoke-Test 'Seleccion explicita entre multiples tablas' {
        Set-DestinationMock -Columns @(New-TestColumns) -TableCount 2
        Assert-Throws { Get-ExcelDestination $script:SharedUrl '' } 'Selecciono tabla arbitraria.'
        Assert-Equal (Get-ExcelDestination $script:SharedUrl 'Inventario1').TableName 'Inventario1' 'Tabla indicada ignorada.'
    }
    Invoke-Test 'Paginacion combina paginas y restringe host' {
        $script:Calls.Clear()
        $script:Handler = {
            param($Method, $Uri, $Body)
            if ($Uri -match '\?page=2$') { return [pscustomobject]@{ value = @([pscustomobject]@{ name = 'segundo' }) } }
            [pscustomobject]@{ value = @([pscustomobject]@{ name = 'primero' }); '@odata.nextLink' = 'https://graph.microsoft.com/v1.0/test?page=2' }
        }
        $items = @(Get-GraphCollection 'https://graph.microsoft.com/v1.0/test')
        Assert-Equal ($items.name -join ',') 'primero,segundo' 'Paginacion incompleta.'
        foreach ($badUrl in @('https://evil.invalid/v1.0/test','http://graph.microsoft.com/v1.0/test',
            'https://graph.microsoft.com.evil.invalid/v1.0/test','https://graph.microsoft.com@evil.invalid/v1.0/test')) {
            $script:BadNext = $badUrl
            $script:Calls.Clear()
            $script:Handler = { [pscustomobject]@{ value = @(); '@odata.nextLink' = $script:BadNext } }
            Assert-Throws { Get-GraphCollection 'https://graph.microsoft.com/v1.0/test' } 'Siguio nextLink ajeno.'
            Assert-Equal $script:Calls.Count 1 'Llamo al host ajeno.'
        }
    }
    Invoke-Test 'SoloLocal no autentica y ComprobarDestino no inserta' {
        $fixture = Join-Path $script:TempRoot 'registro.json'
        New-TestRecord | ConvertTo-Json | Set-Content -LiteralPath $fixture -Encoding UTF8
        $before = $script:Connections
        $result = Invoke-InventoryMain -RegistroJson $fixture -DirectorioSalida (Join-Path $script:TempRoot 'local') -SoloLocal
        Assert-Equal $result.Estado 'SoloLocal' 'Estado local incorrecto.'
        Assert-Equal $script:Connections $before 'SoloLocal inicio sesion.'
        Set-DestinationMock -Columns @(New-TestColumns)
        $script:Contexts.Clear()
        $result = Invoke-InventoryMain -ExcelUrl $script:SharedUrl -ComprobarDestino
        Assert-Equal $result.Estado 'DestinoVerificado' 'No verifico destino.'
        Assert-Equal @($script:Calls | Where-Object { $_.Method -eq 'POST' }).Count 0 'ComprobarDestino inserto.'
    }
    Invoke-Test 'Envia una sola vez y acepta indice cero' {
        Set-DestinationMock -Columns @(New-TestColumns)
        $script:Contexts.Clear()
        $before = $script:Disconnections
        $result = Invoke-InventoryMain -RegistroJson (Join-Path $script:TempRoot 'registro.json') -ExcelUrl $script:SharedUrl -DirectorioSalida (Join-Path $script:TempRoot 'envio')
        Assert-Equal $result.Estado 'Enviado' 'No confirmo insercion.'
        Assert-Equal $result.IndiceFila 0 'Trato indice cero como fallo.'
        Assert-Equal @($script:Calls | Where-Object { $_.Method -eq 'POST' }).Count 1 'Ejecuto mas de un POST.'
        Assert-Equal $script:Contexts[0].MaxRetry 0 'No desactivo reintentos SDK.'
        Assert-Equal $script:Contexts[$script:Contexts.Count - 1].MaxRetry 3 'No restauro reintentos.'
        Assert-Equal $script:Contexts[$script:Contexts.Count - 1].ClientTimeout 300 'No restauro timeout.'
        Assert-Equal $script:Disconnections ($before + 1) 'No cerro Graph.'
    }
    Invoke-Test 'Fallo POST no reintenta, restaura contexto y conserva respaldo' {
        $directory = Join-Path $script:TempRoot 'fallo'
        Set-DestinationMock -Columns @(New-TestColumns) -FailPost
        $script:Contexts.Clear()
        $before = $script:Disconnections
        Assert-Throws { Invoke-InventoryMain -RegistroJson (Join-Path $script:TempRoot 'registro.json') -ExcelUrl $script:SharedUrl -DirectorioSalida $directory -WarningAction SilentlyContinue } 'Oculto fallo Graph.'
        Assert-Equal @($script:Calls | Where-Object { $_.Method -eq 'POST' }).Count 1 'Reintento POST desconocido.'
        Assert-Equal @(Get-ChildItem -LiteralPath $directory -Filter '*.json').Count 1 'Perdio respaldo.'
        Assert-Equal $script:Contexts[$script:Contexts.Count - 1].MaxRetry 3 'No restauro contexto.'
        Assert-Equal $script:Disconnections ($before + 1) 'No cerro Graph despues de fallo.'
    }
    Invoke-Test 'Respuesta sin confirmacion tampoco reintenta' {
        Set-DestinationMock -Columns @(New-TestColumns) -UnconfirmedPost
        $script:Contexts.Clear()
        Assert-Throws { Invoke-InventoryMain -RegistroJson (Join-Path $script:TempRoot 'registro.json') -ExcelUrl $script:SharedUrl -DirectorioSalida (Join-Path $script:TempRoot 'sin-confirmacion') -WarningAction SilentlyContinue } 'Acepto envio sin confirmacion.'
        Assert-Equal @($script:Calls | Where-Object { $_.Method -eq 'POST' }).Count 1 'Reintento sin confirmacion.'
    }
    # Las consultas CIM tambien se simulan: la suite no recopila datos reales.
    $script:SparseCim = $false
    function Get-CimData {
        param([string]$ClassName, [string]$Filter)
        if ($script:SparseCim) {
            switch ($ClassName) {
                'Win32_ComputerSystem' { [pscustomobject]@{ Name=$null; UserName=$null; Manufacturer=$null; Model=$null; TotalPhysicalMemory=$null } }
                'Win32_OperatingSystem' { [pscustomobject]@{ Caption=$null; Version=$null } }
            }
            return
        }
        switch ($ClassName) {
            'Win32_ComputerSystem' { [pscustomobject]@{ Name='PC-TEST'; UserName='DOMINIO\persona'; Manufacturer='Marca'; Model='Modelo'; TotalPhysicalMemory=16GB } }
            'Win32_OperatingSystem' { [pscustomobject]@{ Caption='Windows Test'; Version='10.0' } }
            'Win32_BIOS' { [pscustomobject]@{ SerialNumber='SERIE-1' } }
            'Win32_Processor' { [pscustomobject]@{ Name='CPU Test' } }
            'Win32_SystemEnclosure' { [pscustomobject]@{ ChassisTypes=@(9) } }
            'Win32_LogicalDisk' {
                [pscustomobject]@{ Size=128GB; FreeSpace=32GB }
                [pscustomobject]@{ Size=256GB; FreeSpace=128GB }
            }
            'Win32_NetworkAdapterConfiguration' { [pscustomobject]@{ IPAddress=@('192.168.1.10','169.254.1.1','127.0.0.1','fe80::1') } }
            'SoftwareLicensingProduct' { [pscustomobject]@{ LicenseStatus=1 } }
        }
    }
    function Get-PhysicalDisk {
        [CmdletBinding()] param()
        [pscustomobject]@{ BusType='SATA'; MediaType='SSD' }
        [pscustomobject]@{ BusType='USB'; MediaType='HDD' }
    }
    function Get-ItemProperty {
        [CmdletBinding()] param([string]$Path)
        [pscustomobject]@{ DisplayVersion='TEST'; ReleaseId='OLD' }
    }
    Invoke-Test 'Recopilacion CIM simulada suma discos, filtra red y recorta usuario' {
        $record = Get-InventoryRecord ' TI '
        Assert-InventoryRecord $record
        Assert-Equal $record.Equipo 'PC-TEST' 'Equipo incorrecto.'
        Assert-Equal $record.Usuario 'persona' 'No recorto dominio.'
        Assert-Equal $record.TipoEquipo 'Laptop' 'Tipo equipo incorrecto.'
        Assert-Equal $record.RAM_GB 16 'RAM incorrecta.'
        Assert-Equal $record.DiscoTotal_GB 384 'No sumo discos.'
        Assert-Equal $record.DiscoUsado_GB 224 'Espacio usado incorrecto.'
        Assert-Equal $record.DiscoLibre_GB 160 'Espacio libre incorrecto.'
        Assert-Equal $record.TipoDisco 'SSD' 'Incluyo USB.'
        Assert-Equal $record.IP '192.168.1.10' 'No filtro IP.'
        Assert-Equal $record.Version 'TEST' 'No utilizo DisplayVersion.'
        Assert-Equal $record.Area 'TI' 'No recorto Area.'
    }
    Invoke-Test 'Consultas opcionales ausentes conservan 18 campos sin perder nulls' {
        $script:SparseCim = $true
        try {
            $record = Get-InventoryRecord 'TI'
            Assert-InventoryRecord $record
            Assert-Equal $record.Usuario 'Sin sesion interactiva' 'Usuario ausente mal representado.'
            Assert-Equal $record.Serie 'Desconocido' 'BIOS ausente mal representada.'
            Assert-Equal $record.Procesador 'Desconocido' 'CPU ausente mal representada.'
            Assert-Equal $record.TipoEquipo 'Desconocido' 'Chasis ausente mal representado.'
            Assert-Equal $record.IP 'Desconocido' 'IP ausente mal representada.'
            Assert-Equal $record.Activado 'Desconocido' 'Licencia ausente mal representada.'
            Assert-Equal $record.DiscoTotal_GB $null 'Discos ausentes deben conservar null.'
            Assert-Equal $record.RAM_GB $null 'RAM ausente debe conservar null.'
            $body = New-InventoryRowBody $record @(New-TestColumns) | ConvertFrom-Json
            Assert-Equal $body.values[0].Count 18 'Los nulls desplazaron las columnas.'
        } finally { $script:SparseCim = $false }
    }
} finally {
    # Solo se elimina el directorio temporal unico creado por esta prueba.
    $resolvedTemp = [IO.Path]::GetFullPath($script:TempRoot)
    $tempParent = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if ($resolvedTemp.StartsWith($tempParent, [StringComparison]::OrdinalIgnoreCase) -and
        [IO.Path]::GetFileName($resolvedTemp) -match '^InventarioIT-tests-[a-f0-9]{32}$') {
        Remove-Item -LiteralPath $resolvedTemp -Recurse -Force
    }
}
Write-Host ("Pruebas correctas: {0}; fallidas: {1}." -f $script:Passed, $script:Failures.Count)
if ($script:Failures.Count) { throw ($script:Failures -join [Environment]::NewLine) }
