# Builds index.html (a full HTML document) from app.template.html (the Claude Artifact fragment).
$dir=Split-Path -Parent $MyInvocation.MyCommand.Path
$u=New-Object Text.UTF8Encoding $false
$frag=[IO.File]::ReadAllText("$dir\app.template.html",$u)
# index.html loads the QR Code and Excel libraries from the vendor folder (no internet needed).
# app.template.html keeps the CDN addresses because the Claude Artifact version has no vendor folder.
$frag=$frag.Replace('https://cdnjs.cloudflare.com/ajax/libs/qrcodejs/1.0.0/qrcode.min.js','vendor/qrcode.min.js').Replace('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js','vendor/xlsx.full.min.js')
$head='<!doctype html>'+"`n"+'<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>[hidden]{display:none!important}img{max-width:100%}</style><script src="config.js"></script></head><body>'+"`n"
[IO.File]::WriteAllText("$dir\index.html",($head+$frag+"</body></html>`n"),$u)
"built index.html " + (Get-Item "$dir\index.html").Length + " bytes"
