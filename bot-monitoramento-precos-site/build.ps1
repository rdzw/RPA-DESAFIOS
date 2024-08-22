$exclude = @("venv", "bot-monitoramento-precos-site.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "bot-monitoramento-precos-site.zip" -Force