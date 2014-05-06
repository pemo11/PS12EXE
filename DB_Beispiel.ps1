<#
 .Synopsis
 Datenbankzugriff auf einen MS SQL Server
#>

Add-Type -Assembly System.Data

Set-StrictMode -Version 2.0

<#
 .Synopsis
 Anlegen der Datenbank
#>
function Create-DBDatabase
{
    [CmdletBinding()]
    param([String]$DBName, [String]$ConnectionString)
    try
    {
        $Cn = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $ConnectionString
        $Cn.Open()
        Write-Verbose "Verbindung steht..."
        try
        {
          $Cmd = $Cn.CreateCommand()
          $Sql = "If  Exists (Select name From master.dbo.sysdatabases Where name = N'$DBName') Drop Database [$DBName]"
          $Cmd.CommandText = $Sql
          $Result = $Cmd.ExecuteNonQuery()
          Write-Verbose "Datenbank $DBName wurde gelöscht..."
        }
        catch
        {
          Write-Warning "Fehler beim Löschen der Datenbank $DBName ($_)"
        }
        try
        {
          $Cmd.CommandText = "Create Database $DBName"
          $Result = $Cmd.ExecuteNonQuery()
          Write-Verbose "Datenbank wurde angelegt..."
        }
        catch
        {
          Write-Warning "Fehler beim Anlegen der Datenbank $DBName - Verarbeitung wird abgebrochen ($_)"
          break
        }
    }
    catch 
    {
        Write-Warning "Fehler in Create-DBDatabase ($_)"
    }
    finally
    {
       if ($Cn -ne $Null)
       {
        $Cn.Close()
        $Cn.Dispose()
        $Cn = $Null
        Write-Verbose "Verbindung wurde geschlossen..."
       }
    }
}

Set-Alias -Name DBAnlegen -Value Create-DBDatabase

<#
 .Synopsis
 Anlegen einer Tabelle
#>
function Create-DBTable
{
    [CmdletBinding()]
    param([String]$TableName, [String]$ConnectionString)
    try
    {
        $Cn = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $ConnectionString
        $Cn.Open()
        Write-Verbose "Verbindung steht..."
        # Tablelle vorher löschen, sofern vorhanden
        $Sql = "IF OBJECT_ID('$TableName', 'U') IS NOT NULL  DROP TABLE $TableName"
        $Cmd = $Cn.CreateCommand()
        $Cmd.CommandText = $Sql
        $Result = $Cmd.ExecuteNonQuery()
        $Sql = "Create Table $TableName ("
        $Sql += "LoginID int Identity(1,1) Primary Key Not Null,"
        $Sql += "UserID int  Not Null,"
        $Sql += "LoginTime DateTime  Not Null)"
        $Cmd = $Cn.CreateCommand()
        $Cmd.CommandText = $Sql
        $Result = $Cmd.ExecuteNonQuery()
        Write-Verbose "Tabelle $TableName wurde angelegt..."
        $Cmd.Dispose()
        $Cmd = $null
    }
    catch
    {
        Write-Warning "Fehler beim Anlegen der Tabelle $TableName ($_) - Verarbeitung wird abgebrochen"
        break
    }
    finally
    {
       if ($Cn -ne $Null)
       {
        $Cn.Close()
        $Cn.Dispose()
        $Cn = $Null
        Write-Verbose "Verbindung wurde geschlossen..."
       }
    }
}

Set-Alias -Name TableAnlegen -Value Create-DBTable

<#
 .Synopsis
 Hinzufügen von Tabellen zu einer Datenbank
#>
function Add-DBContent
{
    [CmdletBinding()]
    param([String]$ConnectionString, [Int]$Anzahl, [String]$TableName)
    try
    {
        $Cn = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $ConnectionString
        $Cn.Open()
        Write-Verbose "Verbindung steht..."
        for($i=0;$i-lt$Anzahl;$i++)
        {
            $UserID = 1..10 | Get-Random
            $Cmd = $Cn.CreateCommand()
            [DateTime]$Zeitpunkt = get-date # -format dd/MM/yyyy
            $Cmd.CommandText = "Insert Into $TableName Values($UserID, @Zeitpunkt)"
            $SqlPara = $Cmd.CreateParameter()
            $SqlPara.DbType = "DateTime"
            $sqlPara.ParameterName = "Zeitpunkt"
            $SqlPara.Value = Get-Date
            $Cmd.Parameters.Add($SqlPara)  | Out-Null
            $Result = $Cmd.ExecuteNonQuery()
        }
    }
    catch 
    {
        Write-Warning "Fehler beim Hinzufügen eines Datensatzes ($_)"
    }
    finally
    {
       if ($Cn -ne $Null)
       {
        $Cn.Close()
        $Cn.Dispose()
        $Cn = $Null
        Write-Verbose "Verbindung wurde geschlossen..."
       }
    }
}

Set-Alias -Name DBFuellen -Value Add-DBContent

<#
 .Synopsis
 Abfragen von Datensätzen
#>
function Get-DBContent
{
    [CmdletBinding()]
    param([String]$ConnectionString, [String]$TableName)
    try
    {
        $Cn = New-Object -TypeName System.Data.SqlClient.SqlConnection -ArgumentList $ConnectionString
        $Cn.Open()
        Write-Verbose "Verbindung steht..."
        $Cmd = $Cn.CreateCommand()
        $Cmd.CommandText = "Select * From $TableName"
        $Dr = $Cmd.ExecuteReader()
        while ($Dr.Read())
        {
          New-Object -TypeName PSObject -Property @{LoginID=$Dr.GetInt32($Dr.GetOrdinal("LoginID"));
                                                    UserID=$Dr.GetInt32($Dr.GetOrdinal("UserID"));
                                                    LoginTime=$Dr.GetDateTime(($Dr.GetOrdinal("LoginTime")))}

        }
    }
    catch 
    {
        Write-Warning "Fehler bei der Datenbankabfrage ($_)"
    }
    finally
    {
       if ($Cn -ne $Null)
       {
        $Cn.Close()
        $Cn.Dispose()
        $Cn = $Null
        Write-Verbose "Verbindung wurde geschlossen..."
       }
    }
}

$DBName = "LoginsDB"
$TableName = "Logins"
$CnSt = "Server=.\SQLExpress;Integrated Security=SSPI"

Set-Alias -Name DBAuslesen -Value Get-DBContent

DBAnlegen -DBName $DBName -ConnectionString $CnSt -Verbose

$CnSt = "Server=.\SQLExpress;Initial Catalog=$DBName;Integrated Security=SSPI"

TableAnlegen -TableName $TableName -ConnectionString $CnSt -Verbose

DBFuellen -Anzahl 100 -TableName $TableName -ConnectionString $CnSt -Verbose

DBAuslesen -TableName $TableName -ConnectionString $CnSt -Verbose

