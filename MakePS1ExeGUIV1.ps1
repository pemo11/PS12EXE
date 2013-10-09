<#
,Synopsis
Ps12EXE - PS1-Dateien in Exe einbetten GUI-Version
.Description
Erzeugt eine Exe-Datei mit einer Ps1-Datei als eingebetteter Ressource, die nach dem Start ausgeführt wird
Version 1.0
.Notes
Original-Autor und Idee: Keith Hill
.Notes
Aktuell funktioniert das Skript nur unter der Powershell 2.0
.Notes
Powershell.exe muss mit -STA gestartet werden
#>

Set-StrictMode -Version 2.0

<#
.Synopsis
Bettet eine Ps1-Datei in eine Exe-Datei ein
.Parameter PS1Path
Der Pfad der Ps1-Datei
.Parameter OutputAssemblyPath
Der Pfad der Exe-Datei, die angelegt werden soll
#>
function Make-Exe
{
    [CmdletBinding()]

    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateNotNullOrEmpty()]
        [String]
        $PS1Path,
    
        [Parameter(Mandatory = $true, Position = 1)]
        [String]
        $OutputAssemblyPath
    )

    # $PS1PathCS = $PS1Path -replace "\\", "\\"
    $PS1Name = Split-Path -Path $PS1Path -Leaf

    $AssemblyCode = @"
    using System;
    using System.Collections.Generic;
    using System.Collections.ObjectModel;
    using System.Globalization;
    using System.IO;
    using System.IO.Compression;
    using System.Management.Automation;
    using System.Management.Automation.Host;
    using System.Management.Automation.Runspaces;
    using System.Reflection;
    using System.Security;
    using System.Text;
    using System.Threading;
    using System.Diagnostics;

    namespace PS1ToExeTemplate
    {
        class Program
        {
            private static object _powerShellLock = new object();
            private static PowerShell _powerShellEngine;
            private static readonly PS1Host _psHost = new PS1Host();

            static void Main(string[] args)
            {
            string script = GetScript();
            RunScript(script, args, null);
            }

            private static string GetScript()
            {
            string script = String.Empty;

            Assembly ass = Assembly.GetExecutingAssembly();
            try
            {
                using (Stream st = ass.GetManifestResourceStream("$PS1Name"))
                {
                    StreamReader sr = new StreamReader(st);
                    script = sr.ReadToEnd();
                }
            }
            catch
            {
                script = "Write-Host -Fore Red -Back White";
            }
                return script;
            }

            private static void RunScript(string script, string[] args, object input)
            {
                Debug.WriteLine("Skript:\n " + script);
            lock (_powerShellLock)
            {
                _powerShellEngine = PowerShell.Create();
            }

            try
            {
                _powerShellEngine.Runspace = RunspaceFactory.CreateRunspace(_psHost);
                _powerShellEngine.Runspace.ApartmentState = System.Threading.ApartmentState.STA;
                _powerShellEngine.Runspace.Open();
                _powerShellEngine.AddScript(script);
                _powerShellEngine.AddCommand("Out-Default");
                _powerShellEngine.Commands.Commands[0].MergeMyResults(PipelineResultTypes.Error, PipelineResultTypes.Output);

                if (input != null)
                {
                    Collection<PSObject> results = _powerShellEngine.Invoke(new[] { input });
                    Debug.WriteLine(String.Format("Fertig - {0} Ergebnisse.", results.Count));
                }
                else
                {
                    Collection<PSObject> results = _powerShellEngine.Invoke();
                    Debug.WriteLine(String.Format("Fertig - {0} Ergebnisse.", results.Count));
                }
            }
            catch (SystemException ex)
            {
                Debug.WriteLine("Fehler: " + ex.Message);
            }
            finally
            {
                lock (_powerShellLock)
                {
                    _powerShellEngine.Dispose();
                    _powerShellEngine = null;
                }
            }
        }
    }

    // Die Klasse PS1Host
    class PS1Host: PSHost
    {
        private PSHostUserInterface _psHostUserInterface = new HostUserInterface();

        public override void SetShouldExit(int exitCode)
        {
            Environment.Exit(exitCode);
        }

        public override void EnterNestedPrompt()
        {
            throw new NotImplementedException();
        }

        public override void ExitNestedPrompt()
        {
            throw new NotImplementedException();
        }

        public override void NotifyBeginApplication()
        {
            throw new NotImplementedException();
        }

        public override void NotifyEndApplication()
        {
            throw new NotImplementedException();
        }

        public override string Name
        {
            get { return "PS1ToExeHost"; }
        }

        public override Version Version
        {
            get { return new Version(1, 0); }
        }

        public override Guid InstanceId
        {
            get { return new Guid("E4673B42-84B6-4C43-9589-95FAB8E00EB2"); }
        }
        
        public override PSHostUserInterface UI
        {
            get { return _psHostUserInterface; }
        }

        public override CultureInfo CurrentCulture
        {
            get { return Thread.CurrentThread.CurrentCulture; }
        }

        public override CultureInfo CurrentUICulture
        {
            get { return Thread.CurrentThread.CurrentUICulture; }
        }
    }

    // Ende der Host-Klasse

    // Klasse HostUserInterface

    class HostUserInterface : PSHostUserInterface
    {
        private PSHostRawUserInterface _psRawUserInterface = new HostRawUserInterface();

        public override PSHostRawUserInterface RawUI
        {
            get { return _psRawUserInterface; }
        }

        public override string ReadLine()
        {
            return Console.ReadLine();
        }

        public override SecureString ReadLineAsSecureString()
        {
            throw new NotImplementedException();
        }

        public override void Write(string value)
        {
            string output = value ?? "null";
            Console.Write(output);
        }

        public override void Write(ConsoleColor foregroundColor, ConsoleColor backgroundColor, string value)
        {
            string output = value ?? "null";
            var origFgColor = Console.ForegroundColor;
            var origBgColor = Console.BackgroundColor;
            Console.ForegroundColor = foregroundColor;
            Console.BackgroundColor = backgroundColor;
            Console.Write(output);
            Console.ForegroundColor = origFgColor;
            Console.BackgroundColor = origBgColor;
        }

        public override void WriteLine(string value)
        {
            string output = value ?? "null";
            Console.WriteLine(output);
        }

        public override void WriteErrorLine(string value)
        {
            string output = value ?? "null";
            var origFgColor = Console.ForegroundColor;
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine(output);
            Console.ForegroundColor = origFgColor;
        }

        public override void WriteDebugLine(string message)
        {
            WriteYellowAnnotatedLine(message, "DEBUG");
        }

        public override void WriteVerboseLine(string message)
        {
            WriteYellowAnnotatedLine(message, "VERBOSE");
        }

        public override void WriteWarningLine(string message)
        {
            WriteYellowAnnotatedLine(message, "WARNING");
        }

        private void WriteYellowAnnotatedLine(string message, string annotation)
        {
            string output = message ?? "null";
            var origFgColor = Console.ForegroundColor;
            var origBgColor = Console.BackgroundColor;
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine(String.Format(CultureInfo.CurrentCulture, "{0}: {1}", annotation, output));
            Console.ForegroundColor = origFgColor;
            Console.BackgroundColor = origBgColor;
        }

        public override void WriteProgress(long sourceId, ProgressRecord record)
        {
            throw new NotImplementedException();
        }

        public override Dictionary<string, PSObject> Prompt(string caption, string message, Collection<FieldDescription> descriptions)
        {
            throw new NotImplementedException();
        }

        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName)
        {
            throw new NotImplementedException();
        }

        public override PSCredential PromptForCredential(string caption, string message, string userName, string targetName, PSCredentialTypes allowedCredentialTypes, PSCredentialUIOptions options)
        {
            throw new NotImplementedException();
        }

        public override int PromptForChoice(string caption, string message, Collection<ChoiceDescription> choices, int defaultChoice)
        {
            throw new NotImplementedException();
        }
    }

    // Klasse HostRawUserInterface

    class HostRawUserInterface : PSHostRawUserInterface
    {
        public override KeyInfo ReadKey(ReadKeyOptions options)
        {
            throw new NotImplementedException();
        }

        public override void FlushInputBuffer()
        {
            throw new NotImplementedException();
        }

        public override void SetBufferContents(Coordinates origin, BufferCell[,] contents)
        {
            throw new NotImplementedException();
        }

        public override void SetBufferContents(Rectangle rectangle, BufferCell fill)
        {
            throw new NotImplementedException();
        }

        public override BufferCell[,] GetBufferContents(Rectangle rectangle)
        {
            throw new NotImplementedException();
        }

        public override void ScrollBufferContents(Rectangle source, Coordinates destination, Rectangle clip, BufferCell fill)
        {
            throw new NotImplementedException();
        }

        public override ConsoleColor ForegroundColor
        {
            get { return Console.ForegroundColor; }
            set { Console.ForegroundColor = value; }
        }

        public override ConsoleColor BackgroundColor
        {
            get { return Console.BackgroundColor; }
            set { Console.BackgroundColor = value; }
        }

        public override Coordinates CursorPosition
        {
            get { return new Coordinates(Console.CursorLeft, Console.CursorTop); }
            set { Console.SetCursorPosition(value.X, value.Y); }
        }

        public override Coordinates WindowPosition
        {
            get { return new Coordinates(Console.WindowLeft, Console.WindowTop); }
            set { Console.SetWindowPosition(value.X, value.Y); }
        }

        public override int CursorSize
        {
            get { return Console.CursorSize; }
            set { Console.CursorSize = value; }
        }

        public override Size BufferSize
        {
            get { return new Size(Console.BufferWidth, Console.BufferHeight); }
            set { Console.SetBufferSize(value.Width, value.Height); }
        }

        public override Size WindowSize
        {
            get { return new Size(Console.WindowWidth, Console.WindowHeight); }
            set { Console.SetWindowSize(value.Width, value.Height); }
        }

        public override Size MaxWindowSize
        {
            get { return new Size(Console.LargestWindowWidth, Console.LargestWindowHeight); }
        }

        public override Size MaxPhysicalWindowSize
        {
            get { return new Size(Console.LargestWindowWidth, Console.LargestWindowHeight); }
        }

        public override bool KeyAvailable
        {
            get { return Console.KeyAvailable; }
        }

        public override string WindowTitle
        {
            get { return Console.Title; }
            set { Console.Title = value; }
        }
    }
}
"@

      # Hier beginnt die Function
    $ReferenceAssemblies = "System.dll","System.Data.dll", ([PSObject].Assembly.Location)

    $ErrorLogPfad = Join-Path -Path (Split-Path -Path $PS1Path) -ChildPath PS1ExeErrors.log
    $Cp = New-Object -TypeName System.CodeDom.Compiler.CompilerParameters -ArgumentList $ReferenceAssemblies,$OutputAssemblyPath,$true
    $Cp.TempFiles = New-Object -TypeName System.CodeDom.Compiler.TempFileCollection -ArgumentList ([IO.Path]::GetTempPath())
    $Cp.GenerateExecutable = $true
    $Cp.GenerateInMemory = $false
    $Cp.IncludeDebugInformation = $false
    [void]$Cp.EmbeddedResources.Add($PS1Path)
    $Dict = New-Object -Typename 'System.Collections.Generic.Dictionary[string, string]'
    $Dict.Add("CompilerVersion","v3.5")
    $Provider = New-Object -TypeName Microsoft.CSharp.CSharpCodeProvider -ArgumentList $Dict
    $Results = $Provider.CompileAssemblyFromSource($Cp, $AssemblyCode)
    $StatusListBox.Items.Add("Exe-Datei wurde erstellt")
    $StatusListBox.Items.Add(("Anzahl Fehler: {0:0}" -f $Results.Errors.Count))
    if ($Results.Errors.Count)
    {
        $ErrorLines = ""
        foreach ($Error in $results.Errors)
        {
            $ErrorMsg = "*** Fehler: " + "`t" + $Error.ErrorText + "(" + $Error.Line + ")"
            $StatusListBox.Items.Add($ErrorMsg)
            Add-Content -Path $ErrorLogPfad -Value $ErrorMsg
         }
    }
    else
    {
        $StatusListBox.Items.Add("*** Erfolg - Exe-Datei wurde fehlerfrei erstellt.")
    }
}

<#
.Synopsis
Zeigt das Fenster an
#>
function Show-MainWindow
{
    $MainForm = New-Object -Typename System.Windows.Forms.Form
    $MainStatusBar = New-Object -Typename System.Windows.Forms.StatusBar
    $GroupBox1 = New-Object -Typename System.Windows.Forms.GroupBox
    $GroupBox2 = New-Object -Typename System.Windows.Forms.GroupBox
    $StatusListBox = New-Object -Typename System.Windows.Forms.ListBox
    $ExeErstellenButton = New-Object -Typename System.Windows.Forms.Button
    $PS1AuswahlButton = New-Object -Typename System.Windows.Forms.Button
    $InitialFormWindowState = New-Object -Typename System.Windows.Forms.FormWindowState
    $ExePfadTextBox = New-Object -Typename System.Windows.Forms.TextBox
    $Label1 = New-Object -Typename System.Windows.Forms.Label

    # Skriptblock für das Anklicken des PS1Auswahl-Buttons
    $PS1AuswahlButton_OnClick=
    {
      $Ofd = New-Object -TypeName System.Windows.Forms.OpenFileDialog
      $Ofd.Title = "Auswahl PS1-Datei"
      $Ofd.Filter = "PS1-Dateien (*.ps1)|*.ps1|Alle Dateien|*.*"
      if ($Ofd.ShowDialog() -eq "OK")
      {
        $Script:Ps1Pfad = $Ofd.FileName
        $StatusListBox.Items.Add($Ps1Pfad + " wurde gewählt")
        $ExeErstellenButton.Enabled = $true;
      }
    }

    # Skriptblock für das Anklicken des Exe-Erstellen-Buttons
    $ExeErstellenButton_OnClick=
    {
        $AssPfad = $ExePfadTextBox.Text
        # Gibt es den Pfad?
        if (!(Test-Path -Path $AssPfad))
        {
            # Ist es ein gültiger Pfad?
            if (Test-Path -Path $AssPfad -IsValid)
            {
              # Soll das Verzeichnis angelegt werden?
              <#
Choice-Dialog wird in der Befehlszeile auch in der Befehlszeile angezeigt
$JaChoice = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription "Ja", "Verzeichnis anlegen"
$NeinChoice = New-Object -TypeName System.Management.Automation.Host.ChoiceDescription "Nein", "Verzeichnis nicht anlegen"
if ($Host.UI.PromptForChoice("Wie soll es weitergehen?", "Verzeichnis $AssPfad anlegen?",
@($JaChoice, $NeinChoice), 0) -eq 0)
#>
               if ([System.Windows.MessageBox]::Show("Wie soll es weitergehen?", "Verzeichnis $AssPfad anlegen?","YesNo"))
               {
                 md -Path $AssPfad -ErrorAction Ignore
                 if (!$?)
                 {
                    $StatusListBox.Items.Add("$AssPfad konnte nicht angelegt werden - Skript wird beendet.")
                    Exit -2
                 }
                 else
                 {
                    $StatusListBox.Items.Add("Das Verzeichnis $AssPfad wurde angelegt.")
                 }
               }
             }
             else
             {
                $StatusListBox.Items.Add("Ungültiges Verzeichnis $AssPfad - Skript wird beendet.")
                Exit -3
             }
        }
        $StatusListBox.Items.Add($Ps1Pfad + " wird in Exe-Datei konvertiert.")
        $AssName = [System.IO.Path]::GetFileNameWithoutExtension($Ps1Pfad)
        $AssName += "_Konvertiert.exe"
        $AssPfad = Join-Path -Path $AssPfad -ChildPath $AssName
        Make-Exe -PS1Path $Ps1Pfad -OutputAssemblyPath $AssPfad
        $ExeErstellenButton.Enabled = $false
    }

    # Skritblock für das Laden des Fensters
    $OnLoadForm=
    {
  $MainForm.WindowState = $InitialFormWindowState
        $ExeErstellenButton.Enabled = $false
        $ExePfadTextBox.Text = $env:USERPROFILE + "\PS1Exe"
    }

    $MainForm.ClientSize = New-Object -Typename System.Drawing.Size -ArgumentList 500, 440
    $MainForm.Font = New-Object -Typename System.Drawing.Font("Arial Narrow",12,1,3,0)
    $MainForm.Name = "Mainform"
    $MainForm.Text = "PS1 in Exe einbetten"

    $MainStatusBar.Location = New-Object -Typename System.Drawing.Point -ArgumentList 300, 0
    $MainStatusBar.Name = "MainStatusBar"
    $MainStatusBar.Size = New-Object -Typename System.Drawing.Size -ArgumentList 414, 22
    $MainStatusBar.TabIndex = 3
    $MainStatusBar.Text = "OK"

    $MainForm.Controls.Add($MainStatusBar)

    $GroupBox1.Location = New-Object -Typename System.Drawing.Point -ArgumentList 32, 200
    $GroupBox1.Name = "groupBox1"
    $GroupBox1.Size = New-Object -Typename System.Drawing.Size -ArgumentList 400,200
    $GroupBox1.TabStop = $False
    $GroupBox1.Text = "Status"

    $MainForm.Controls.Add($GroupBox1)

    $ExePfadTextBox.BackColor = [System.Drawing.Color]::FromArgb(255,192,255,192)
    $ExePfadTextBox.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 18, 70
    $ExePfadTextBox.Name = "ExePfadTextBox"
    $ExePfadTextBox.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 364, 26

    $GroupBox2.Controls.Add($ExePfadTextBox)

    $Label1.Location = New-Object -TypeName System.Drawing.Point -ArgumentList 16, 44
    $Label1.Name = "label1"
    $Label1.Size = New-Object -TypeName System.Drawing.Size -ArgumentList 88,24
    $Label1.TabIndex = 1
    $Label1.Text = "Exe-Pfad:"

    $Groupbox2.Controls.Add($Label1)

    $GroupBox2.Location = New-Object -Typename System.Drawing.Point -ArgumentList 32, 72
    $GroupBox2.Name = "GroupBox2"
    $GroupBox2.Size = New-Object -Typename System.Drawing.Size -ArgumentList 400,120
    $GroupBox2.TabStop = $False
    $GroupBox2.Text = "Exe-Datei erstellen"

    $MainForm.Controls.Add($GroupBox2)

    $StatusListBox.ItemHeight = 20
    $StatusListBox.Location = New-Object -Typename System.Drawing.Point -ArgumentList 8, 26
    $StatusListBox.Size = New-Object -Typename System.Drawing.Size -ArgumentList 380, 164
    $StatusListBox.HorizontalScrollbar = $true
    $StatusListBox.TabIndex = 0

    $GroupBox1.Controls.Add($StatusListBox)

    $ExeErstellenButton.Location = New-Object -Typename System.Drawing.Point -ArgumentList 220, 18
    $ExeErstellenButton.Name = "ExeErstellenButton"
    $ExeErstellenButton.Size = New-Object -Typename System.Drawing.Size -ArgumentList 166, 46
    $ExeErstellenButton.TabIndex = 1
    $ExeErstellenButton.Text = "Exe-Erstellen"
    $ExeErstellenButton.add_Click($ExeErstellenButton_OnClick)

    $GroupBox2.Controls.Add($ExeErstellenButton)

    $PS1AuswahlButton.Location = New-Object -Typename System.Drawing.Point -ArgumentList 28, 20
    $PS1AuswahlButton.Name = "PS1AuswahlButton"
    $PS1AuswahlButton.Size = New-Object -Typename System.Drawing.Size -ArgumentList 166, 46
    $PS1AuswahlButton.TabIndex = 0
    $PS1AuswahlButton.Text = "Ps1-Auswahl"
    $PS1AuswahlButton.add_Click($PS1AuswahlButton_OnClick)

    $MainForm.Controls.Add($PS1AuswahlButton)

    $InitialFormWindowState = $MainForm.WindowState
    $MainForm.add_Load($OnLoadForm)
    [System.Windows.Forms.Application]::Run($MainForm)
}

# Assemblies nachladen
Add-Type -AssemblyName System.Windows.Forms, System.Drawing

# Fenster anzeigen
Show-MainWindow 