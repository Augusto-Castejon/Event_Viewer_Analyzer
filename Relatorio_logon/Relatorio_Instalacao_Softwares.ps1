# Caminho do relatório de instalações e atualizações de software
$outputPath = "C:\Temp\Relatorio_logon\Relatorio_Instalacao_Softwares.html"
$startTime = (Get-Date).AddDays(-7)
$hostname = $env:COMPUTERNAME

# Eventos de instalação: MsiInstaller ID 1033 e 11707 no log Application
$msiEvents = Get-WinEvent -FilterHashtable @{
    LogName = 'Application';
    ProviderName = 'MsiInstaller';
    Id = 1033, 11707;
    StartTime = $startTime
} | ForEach-Object {
    $xml = [xml]$_.ToXml()
    $time = $_.TimeCreated
    $sid = $_.UserId
    try {
        $user = ([System.Security.Principal.SecurityIdentifier]$sid).Translate([System.Security.Principal.NTAccount]).Value
    } catch {
        $user = $sid
    }
    $mensagem = $_.Message

    $software = ($mensagem -split 'Product Name: ')[1] -split '\n' | Select-Object -First 1
    $tipo = if ($mensagem -like '*Updating Product*') { "Atualização" } else { "Instalação" }
    $versao = ($mensagem -split 'Product Version: ')[1] -split '\n' | Select-Object -First 1
    $codigo = ($mensagem -split 'Product Code: ')[1] -split '\n' | Select-Object -First 1

    [PSCustomObject]@{
        DataHora = $time
        Usuario = $user
        Software = $software.Trim()
        Tipo = $tipo
        Versao = $versao.Trim()
        Codigo = $codigo.Trim()
        Mensagem = $mensagem.Trim()
    }
}

# Estatísticas rápidas
$total = $msiEvents.Count
$instalacoes = ($msiEvents | Where-Object { $_.Tipo -eq 'Instalação' }).Count
$atualizacoes = ($msiEvents | Where-Object { $_.Tipo -eq 'Atualização' }).Count

# Gera HTML
$html = @"
<html>
<head>
<title>Relatório de Instalação e Atualização de Softwares</title>
<style>
body { font-family: Arial; margin: 20px; }
h1 { color: #2a3f5f; }
table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
th { background-color: #f2f2f2; }
</style>
</head>
<body>
<img src="images.jpg" alt="Logo" style="height: 80px; margin-bottom: 20px;">
<h1>Relatório de Instalações e Atualizações de Software - Últimos 7 dias</h1>
<h3>Computador: $hostname</h3>
<p><b>Total de eventos:</b> $total<br>
<b>Instalações:</b> $instalacoes<br>
<b>Atualizações:</b> $atualizacoes</p>
<table>
<tr><th>Data/Hora</th><th>Usuário</th><th>Software</th><th>Tipo</th><th>Versão</th><th>Código</th><th>Mensagem</th></tr>
"@

foreach ($event in $msiEvents | Sort-Object DataHora -Descending) {
    $html += "<tr><td>$($event.DataHora)</td><td>$($event.Usuario)</td><td>$($event.Software)</td><td>$($event.Tipo)</td><td>$($event.Versao)</td><td>$($event.Codigo)</td><td>$($event.Mensagem)</td></tr>"
}

$html += "</table></body></html>"
$html | Out-File -Encoding UTF8 $outputPath
