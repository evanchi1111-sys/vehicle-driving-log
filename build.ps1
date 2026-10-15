# Builds index.html (a full HTML document) from app.template.html (the Claude Artifact fragment).
$dir=Split-Path -Parent $MyInvocation.MyCommand.Path
$u=New-Object Text.UTF8Encoding $false
$frag=[IO.File]::ReadAllText("$dir\app.template.html",$u)
$head='<!doctype html>'+"`n"+'<html lang="zh-Hant"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1,viewport-fit=cover"><style>[hidden]{display:none!important}img{max-width:100%}</style></head><body>'+"`n"
[IO.File]::WriteAllText("$dir\index.html",($head+$frag+"</body></html>`n"),$u)
"built index.html " + (Get-Item "$dir\index.html").Length + " bytes"
