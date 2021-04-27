# Restart Windows Update Service
Restart-Service wuauserv -Force

wuauclt.exe /detectnow

wuauclt.exe /reportnow
