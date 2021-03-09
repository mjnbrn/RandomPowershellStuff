$Wordlist = (Invoke-restmethod "https://gist.githubusercontent.com/deekayen/4148741/raw/98d35708fa344717d8eee15d11987de6c8e26d7d/1-1000.txt") -split "`n"
$Word = Get-Random $Wordlist
$SearchResults = docker search --limit 100 $word
$DockerName = (docker search --limit 100 --format "{{json . }}" $word |ConvertFrom-Json |Get-Random  |Select name).name
while ($DockerName -in (Get-Content .\CrazyDocker.txt)){$DockerName = (docker search --limit 100 --format "{{json . }}" --filter is-automated=true $word |ConvertFrom-Json |Get-Random  |Select name).name}
docker pull $DockerName
& docker run $DockerName
