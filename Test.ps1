param([Int]$Anzahl=3)
Write-Host -Fore Red -Back White "Es geht los... "
1..$Anzahl | Foreach-Object {
 "Durchlauf Nr. $_"
}
"Das Ergebnis: $(22/7)"