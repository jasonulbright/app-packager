#Requires -Version 5.1
$ErrorActionPreference = 'Stop'

# Exercise the actual upload function against a loopback-only HTTP receiver.
# The payload is a harmless source fixture; no Azure/Graph access is involved.
$fixture = Join-Path $PSScriptRoot 'Show-ProcessArguments.ps1'
$module = Join-Path $PSScriptRoot '..\Packagers\AppPackagerCommon.psd1'
$portProbe = New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback, 0)
$portProbe.Start()
$port = $portProbe.LocalEndpoint.Port
$portProbe.Stop()
$listener = New-Object Net.HttpListener
$listener.Prefixes.Add("http://127.0.0.1:$port/")
$client = [powershell]::Create()
try {
    $listener.Start()
    [void]$client.AddScript({
        param($ModulePath, $PayloadPath, $UploadUri)
        Import-Module $ModulePath -Force -DisableNameChecking
        Invoke-AzureBlobUpload -Uri $UploadUri -FilePath $PayloadPath
    }).AddArgument($module).AddArgument($fixture).AddArgument("http://127.0.0.1:$port/blob?test=1")
    $pendingClient = $client.BeginInvoke()
    for ($i = 0; $i -lt 2; $i++) {
        $pendingRequest = $listener.BeginGetContext($null, $null)
        if (-not $pendingRequest.AsyncWaitHandle.WaitOne(10000)) {
            throw ('Local receiver timed out: ' + ($client.Streams.Error -join '; '))
        }
        $context = $listener.EndGetContext($pendingRequest)
        $received = New-Object IO.MemoryStream
        try {
            $context.Request.InputStream.CopyTo($received)
            if ($context.Request.QueryString['comp'] -eq 'block') {
                $expectedBytes = [IO.File]::ReadAllBytes($fixture)
                $receivedBytes = $received.ToArray()
                [pscustomobject]@{
                    PowerShell = $PSVersionTable.PSVersion.ToString()
                    ExpectedBytes = $expectedBytes.Length
                    UploadedBytes = $receivedBytes.Length
                    BytesMatch = [Convert]::ToBase64String($expectedBytes) -eq [Convert]::ToBase64String($receivedBytes)
                    UploadedPrefix = [Text.Encoding]::ASCII.GetString($receivedBytes, 0, [Math]::Min(70, $receivedBytes.Length))
                }
            }
            $context.Response.StatusCode = 201
            $context.Response.ContentLength64 = 0
            $context.Response.Close()
        } finally { $received.Dispose() }
    }
    [void]$client.EndInvoke($pendingClient)
    if ($client.HadErrors) { throw ($client.Streams.Error -join '; ') }
} finally {
    $client.Stop()
    $client.Dispose()
    $listener.Close()
}
