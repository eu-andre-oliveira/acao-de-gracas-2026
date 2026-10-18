# Gera um link de download direto para cada foto original que esta sincronizada
# na pasta do OneDrive e grava o campo "original" em galeria/index.html.
# Uso (PowerShell, na raiz do projeto):
#   powershell -ExecutionPolicy Bypass -File scripts\gerar-links.ps1

$ErrorActionPreference = 'Stop'

# Cliente publico do Microsoft Graph PowerShell (login por codigo no navegador)
$ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$Scope    = 'Files.ReadWrite offline_access'

# 1. Localiza a pasta sincronizada e o caminho dela dentro do OneDrive
$base  = Join-Path $env:OneDrive 'Documentos\temp'
$pasta = Get-ChildItem $base -Directory | Where-Object { $_.Name -like 'fotos*2026' } | Select-Object -First 1
if (-not $pasta) { throw "Pasta nao encontrada dentro de $base" }
$rel    = $pasta.FullName.Substring($env:OneDrive.Length).TrimStart('\') -replace '\\', '/'
$relEnc = ($rel.Split('/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'
$nomesLocais = Get-ChildItem $pasta.FullName -File | ForEach-Object { $_.Name }
Write-Host "Pasta: $($pasta.FullName) ($($nomesLocais.Count) arquivos)"

# 2. Login (device code)
$dc = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/consumers/oauth2/v2.0/devicecode' -Body @{ client_id = $ClientId; scope = $Scope }
Write-Host $dc.message
$token = $null
while (-not $token) {
  Start-Sleep -Seconds $dc.interval
  try {
    $t = Invoke-RestMethod -Method Post -Uri 'https://login.microsoftonline.com/consumers/oauth2/v2.0/token' -Body @{
      grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
      client_id   = $ClientId
      device_code = $dc.device_code
    }
    $token = $t.access_token
  } catch {
    $erro = $_.ErrorDetails.Message | ConvertFrom-Json
    if ($erro.error -ne 'authorization_pending') { throw }
  }
}
$h = @{ Authorization = "Bearer $token" }

# 3. Lista os arquivos da pasta no OneDrive (id e nome)
$itens = @()
$url = "https://graph.microsoft.com/v1.0/me/drive/root:/$($relEnc):/children?`$select=id,name&`$top=200"
while ($url) {
  $r = Invoke-RestMethod -Uri $url -Headers $h
  $itens += $r.value
  $url = $r.'@odata.nextLink'
}
$porNome = @{}
foreach ($i in $itens) { $porNome[$i.name] = $i.id }
Write-Host "No OneDrive: $($itens.Count) arquivos"

# 4. Cria o link de visualizacao anonima de cada arquivo (com nova tentativa se o Graph limitar)
function New-ShareLink($id) {
  $body = '{"type":"view","scope":"anonymous"}'
  for ($n = 0; $n -lt 5; $n++) {
    try {
      $r = Invoke-RestMethod -Method Post -Uri "https://graph.microsoft.com/v1.0/me/drive/items/$id/createLink" -Headers $h -ContentType 'application/json' -Body $body
      return $r.link.webUrl
    } catch {
      if ($_.Exception.Response -and [int]$_.Exception.Response.StatusCode -eq 429) { Start-Sleep -Seconds 30 } else { throw }
    }
  }
  throw "Falhou apos varias tentativas: $id"
}

$links = [ordered]@{}
$k = 0
foreach ($nome in $nomesLocais) {
  $k++
  if (-not $porNome.ContainsKey($nome)) { Write-Warning "Nao encontrado no OneDrive: $nome"; continue }
  # O link de compartilhamento abre a foto no visualizador do OneDrive, que tem o botao de baixar.
  # A API de download direto (api.onedrive.com) exige login e nao funciona para links anonimos.
  $links[$nome] = New-ShareLink $porNome[$nome]
  Write-Progress -Activity 'Gerando links' -Status $nome -PercentComplete (100 * $k / $nomesLocais.Count)
  Start-Sleep -Milliseconds 200
}
Write-Progress -Activity 'Gerando links' -Completed

$jsonPath = Join-Path $PSScriptRoot 'links.json'
$links | ConvertTo-Json | Set-Content -Path $jsonPath -Encoding UTF8
Write-Host "Links salvos em $jsonPath ($($links.Count) itens)"

# 5. Atualiza galeria/index.html: adiciona "original" a cada item de fotos
$html  = Join-Path $PSScriptRoot '..\galeria\index.html'
$texto = [IO.File]::ReadAllText($html, [Text.Encoding]::UTF8)
$padrao = '"nome": "([^"]+)"(, "original": "[^"]*")?\}'
$texto = [regex]::Replace($texto, $padrao, {
  param($m)
  $nome = $m.Groups[1].Value
  if ($links.Contains($nome)) { '"nome": "' + $nome + '", "original": "' + $links[$nome] + '"}' } else { $m.Value }
})
[IO.File]::WriteAllText($html, $texto, (New-Object Text.UTF8Encoding $false))
Write-Host 'galeria/index.html atualizado.'
