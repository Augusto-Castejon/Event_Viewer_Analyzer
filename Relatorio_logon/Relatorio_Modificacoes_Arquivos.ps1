# Caminho do relatório de modificações/exclusões
$outputPath = "C:\Temp\Relatorio_logon\Relatorio_Modificacoes_Arquivos.html"
$startTime = (Get-Date).AddDays(-7)

# Verifica se auditoria está ativada
$events = Get-WinEvent -FilterHashtable @{
    LogName = 'Security';
    Id = 4663;
    StartTime = $startTime
} | ForEach-Object {
    $xml = [xml]$_.ToXml()
    $time = $_.TimeCreated
    $user = $xml.Event.EventData.Data | Where-Object { $_.Name -eq 'SubjectUserName' } | Select-Object -ExpandProperty '#text'
    $accessMask = $xml.Event.EventData.Data | Where-Object { $_.Name -eq 'AccessMask' } | Select-Object -ExpandProperty '#text'
    $objName = $xml.Event.EventData.Data | Where-Object { $_.Name -eq 'ObjectName' } | Select-Object -ExpandProperty '#text'
    $objType = $xml.Event.EventData.Data | Where-Object { $_.Name -eq 'ObjectType' } | Select-Object -ExpandProperty '#text'

    # Mapeia a intenção do acesso
    $acao = switch ($accessMask) {
        "0x2" { "Modificar/Escrever" }
        "0x10000" { "Excluir" }
        "0x6" { "Ler/Modificar" }
        "0xc0000" { "Modificar/Escrever" }
        "0x40000" { "Modificar/Escrever" }
        default { "Outro ($accessMask)" }
    }

    [PSCustomObject]@{
        DataHora = $time
        Usuario = $user
        Objeto = $objName
        Tipo = $objType
        Acao = $acao
    }
}

# Gera HTML
$html = @"
<html>
<head>
<title>Relatório de Modificações de Arquivos</title>
<style>
body { font-family: Arial; margin: 20px; }
h1 { color: #2a3f5f; }
table { border-collapse: collapse; width: 100%; }
th, td { border: 1px solid #ccc; padding: 8px; text-align: left; }
th { background-color: #f2f2f2; }
</style>
</head>
<body>
<img src="images.jpg" alt="Logo" style="height: 80px; margin-bottom: 20px;">
<h1>Relatório de Modificações/Exclusões de Arquivos - Últimos 7 dias</h1>
<table>
<tr><th>Data/Hora</th><th>Usuário</th><th>Objeto</th><th>Tipo</th><th>Ação</th></tr>
"@

foreach ($event in $events | Sort-Object DataHora -Descending) {
    $html += "<tr><td>$($event.DataHora)</td><td>$($event.Usuario)</td><td>$($event.Objeto)</td><td>$($event.Tipo)</td><td>$($event.Acao)</td></tr>"
}

$html += "</table></body></html>"
$html | Out-File -Encoding UTF8 $outputPath
