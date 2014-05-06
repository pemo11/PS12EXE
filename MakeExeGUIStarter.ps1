<#
 .Synopsis
 Startet das Skript MakePS1ExeGUIV1.ps1
#>

# Hier den Pfad von MakePS1ExeGUIV1.ps1 eintragen, wenn sich beide nicht 
# im selben Verzeichnis befinden
$SkriptPfad = ".\MakePS1ExeGUIV1.ps1"

PowerShell -Version 2.0 -NoProfile -STA -File $SkriptPfad

# Danach liegt die Ps1-Datei als Exe-Datei mit dem Namen Skriptname_Konvertiert.exe vor

# Wichtig: Die Exe-Datei nicht innerhalb der PowerSHell ISE starten, sondern außerhalb
