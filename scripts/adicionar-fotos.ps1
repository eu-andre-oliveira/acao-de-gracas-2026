# Adiciona a galeria as fotos novas da pasta sincronizada do OneDrive.
# Para cada foto que ainda nao esta em galeria/index.html:
#   1. gera a miniatura em galeria/thumbs/ (lado maior de 800 px, como as existentes)
#   2. cria o link de compartilhamento no OneDrive
#   3. acrescenta o item ao array "fotos" de galeria/index.html
# Uso (PowerShell, na raiz do projeto):
#   powershell -ExecutionPolicy Bypass -File scripts\adicionar-fotos.ps1

$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing

# Cliente publico do Microsoft Graph PowerShell (login por codigo no navegador)
$ClientId = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
$Scope    = 'Files.ReadWrite offline_access'
$Lado     = 800

$raiz  = Join-Path $PSScriptRoot '..'
$html  = Join-Path $raiz 'galeria\index.html'
$thumbs = Join-Path $raiz 'galeria\thumbs'
$texto = [IO.File]::ReadAllText($html, [Text.Encoding]::UTF8)

# 1. Fotos ja cadastradas e fotos novas na pasta sincronizada
$jaCadastradas = @{}
foreach ($m in [regex]::Matches($texto, '"nome": "([^"]+)"')) { $jaCadastradas[$m.Groups[1].Value] = $true }

$base  = Join-Path $env:OneDrive 'Documentos\temp'
$pasta = Get-ChildItem $base -Directory | Where-Object { $_.Name -like 'fotos*2026' } | Select-Object -First 1
if (-not $pasta) { throw "Pasta nao encontrada dentro de $base" }
$rel    = $pasta.FullName.Substring($env:OneDrive.Length).TrimStart('\') -replace '\\', '/'
$relEnc = ($rel.Split('/') | ForEach-Object { [uri]::EscapeDataString($_) }) -join '/'

$novos = Get-ChildItem $pasta.FullName -File | Where-Object { -not $jaCadastradas.ContainsKey($_.Name) } | Sort-Object Name
Write-Host "Pasta: $($pasta.FullName)"
Write-Host "Fotos novas: $($novos.Count)"
if ($novos.Count -eq 0) { Write-Host 'Nada a fazer.'; return }

# 2. Proximo numero de miniatura
$numeros = Get-ChildItem $thumbs -Filter '*.jpg' | ForEach-Object { [int][IO.Path]::GetFileNameWithoutExtension($_.Name) }
$proximo = [int](($numeros | Measure-Object -Maximum).Maximum) + 1

# 3. Login (device code)
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

# 4. Ids dos arquivos no OneDrive
$itens = @()
$url = "https://graph.microsoft.com/v1.0/me/drive/root:/$($relEnc):/children?`$select=id,name&`$top=200"
while ($url) {
  $r = Invoke-RestMethod -Uri $url -Headers $h
  $itens += $r.value
  $url = $r.'@odata.nextLink'
}
$porNome = @{}
foreach ($i in $itens) { $porNome[$i.name] = $i.id }

# Link de visualizacao anonima (com nova tentativa se o Graph limitar)
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

# Miniatura com lado maior de $Lado px, em JPEG qualidade 82
function Save-Thumb($origem, $destino) {
  $img = [System.Drawing.Image]::FromFile($origem)
  try {
    # [double] evita que o PowerShell escolha a sobrecarga de inteiros e arredonde a escala
    $escala = [Math]::Min([double]1, [double]$Lado / [double][Math]::Max($img.Width, $img.Height))
    $w = [int][Math]::Round($img.Width * $escala)
    $h = [int][Math]::Round($img.Height * $escala)
    $bmp = New-Object System.Drawing.Bitmap $w, $h
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
    $g.DrawImage($img, 0, 0, $w, $h)
    $codec = [System.Drawing.Imaging.ImageCodecInfo]::GetImageEncoders() | Where-Object { $_.MimeType -eq 'image/jpeg' }
    $params = New-Object System.Drawing.Imaging.EncoderParameters 1
    $params.Param[0] = New-Object System.Drawing.Imaging.EncoderParameter ([System.Drawing.Imaging.Encoder]::Quality), 82L
    $bmp.Save($destino, $codec, $params)
    $g.Dispose(); $bmp.Dispose()
  } finally { $img.Dispose() }
}

# 5. Processa cada foto nova
$novosItens = @()
foreach ($f in $novos) {
  if (-not $porNome.ContainsKey($f.Name)) { Write-Warning "Nao encontrado no OneDrive (aguarde a sincronizacao): $($f.Name)"; continue }
  $numero = '{0:D4}' -f [int]$proximo
  $arquivoThumb = "$numero.jpg"
  Save-Thumb $f.FullName (Join-Path $thumbs $arquivoThumb)
  $link = New-ShareLink $porNome[$f.Name]
  $novosItens += [ordered]@{ thumb = "thumbs/$arquivoThumb"; nome = $f.Name; original = $link }
  Write-Host "$arquivoThumb <- $($f.Name)"
  $proximo++
  Start-Sleep -Milliseconds 200
}

if ($novosItens.Count -eq 0) { Write-Host 'Nenhuma foto nova foi processada.'; return }

# 6. Acrescenta os itens ao array "fotos" (fim do array)
$trecho = ($novosItens | ForEach-Object { ($_ | ConvertTo-Json -Compress) }) -join ', '
$fim = $texto.LastIndexOf('}];')
if ($fim -lt 0) { throw 'Nao encontrei o fim do array "fotos" em index.html' }
$texto = $texto.Substring(0, $fim + 1) + ', ' + $trecho + $texto.Substring($fim + 1)
[IO.File]::WriteAllText($html, $texto, (New-Object Text.UTF8Encoding $false))
Write-Host "$($novosItens.Count) foto(s) adicionada(s) a galeria/index.html."
