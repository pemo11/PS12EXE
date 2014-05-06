<#
 .Synopsis
 Inhalt der ZF-SQL-Server-Datenbank ausgeben
#>

Add-Type -Assembly System.Data

Set-StrictMode -Version 2.0

<#
 ,Synopsis
 Abfragen von Datensätzen
#>
function Get-BildDaten
{
    try
    {
        Write-Verbose "Get-BildDaten: Verbindung steht..." -Verbose
        $Cmd = $Cn.CreateCommand()
        $Cmd.CommandText = "Select * From Bild"
        $Dr = $Cmd.ExecuteReader()
        while ($Dr.Read())
        {
          New-Object -TypeName PSObject -Property @{Index=$Dr.GetInt32($Dr.GetOrdinal("BildIndex"));
                                                    FileName=$Dr.GetString($Dr.GetOrdinal("Dateiname"));
                                                    KameraCodeIndex=$Dr.GetInt32($Dr.GetOrdinal("KameraCodeIndex"));
                                                   }

        }
    }
    finally
    {
       if ($Cmd -ne $Null)
       {
        $Cmd.Dispose()
        $Cmd = $null
       }
    }
}

$CnSt = "Data Source=.\SQLEXPRESS;Initial Catalog=ZF1570;Integrated Security=SSPI"
$Cn = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $CnSt

try
{
    $Cn.Open()
    Write-Verbose "Datenbankverbindung wurde geöffnet." -Verbose
}
catch
{
    Write-Warning "Fehler beim Öffnen der Datenbankverbindung ($_)"
}

Get-BildDaten 

if ($Cn -ne $null)
{
    try
    {
     $Cn.Close()
     Write-Verbose "Datenbankverbindung wurde geschlossen." -Verbose
    }
    catch
    {
      Write-Warning "Fehler beim Schließen der Datenbankverbindung ($_)"
    }
}