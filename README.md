# DAVINCI Powershell Skripte

Dieses Repository enthält [Powershell-Skripte](https://docs.microsoft.com/de-de/powershell/scripting/getting-started/getting-started-with-windows-powershell) zur Automation von Aufgaben rund um [DAVINCI](https://davinci.stueber.de).

## Was ist DAVINCI?

DAVINCI ist eine Software zum Erstellen und Publizieren von Stundenplänen, Vertretungsplänen, Kursplänen und Prüfungsplänen in Schulen und an Universitäten. 

## Systemvoraussetzungen

Die Systemvoraussetzungen für eine erfolgreiche Ausführung der PowerShell-Skripte ist PowerShell 5. PowerShell 5 ist unter Windows 2016 und Windows 10 bereits vorinstalliert, für ältere Windows-Systeme (Windows 7 Service Pack 1, Windows 8.1, Windows Server 2008 R2, Windows Server 2012, Windows Server 2012 R2) muss das Windows Management Framework 5.1 installiert werden:

1. Installiere die .NET Framework Runtime (4.5 oder höher): https://dotnet.microsoft.com/download

2. Installiere das Windows Management Framework 5.1 (enthält PowerShell 5): https://www.microsoft.com/en-us/download/details.aspx?id=54616

Die Ausführung von Powershell-Skripten unter Windows ist standardmäßig nicht erlaubt, nur die Shell darf interaktiv benutzt werden. Dies kann man als Administrator jedoch ändern:

1. Starte Powershell als Administrator: `Start > Windows Powershell > Windows Powershell`

2. Tippe `Set-ExecutionPolicy -ExecutionPolicy Unrestricted` ein und bestätige.

Mehr Infos zum Cmdlet `Set-ExecutionPolicy` findest Du in der [Microsoft-Dokumentation](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.security/set-executionpolicy?view=powershell-6).

## Die Powershell-Skripte

Eine ausführliche Beschreibung der einzelnen Skripte findet sich im [Wiki](https://github.com/stuebersystems/davinci.powershell/wiki).

## Kann ich mithelfen?

Ja, sehr gerne. Der beste Weg mitzuhelfen ist es, den Quellcode auszuprobieren, Rückmeldung per Issue-Tracker zu geben und/oder eigene Pull-Requests zu generieren.
