[CmdletBinding(DefaultParameterSetName = 'Path')]

Param (
	[Parameter(Mandatory, Position = 0, ParameterSetName = 'Path')]
	[ValidateScript({[IO.File]::Exists($_) -And [IO.Path]::GetExtension($_) -eq '.exe'})]
	[String]$Path,

	[Parameter(Mandatory, Position = 0, ParameterSetName = 'Name')]
	[String]$Name,

	[Parameter(Mandatory, Position = 0, ParameterSetName = 'PID')]
	[Alias('Pid', 'ProcessId')]
	[Int]$Id,

	[Parameter(Mandatory, Position = 0, ParameterSetName = 'Process')]
	[Diagnostics.Process]$Process,

	[Parameter(Mandatory, ValueFromPipeline, ParameterSetName = 'InputObject')]
	[Object]$InputObject,

	[Parameter(Mandatory, ParameterSetName = 'Previous')]
	[Switch]$Previous
)

[String]$TypeDefinition = @(
	'using System;',
	'using System.Runtime.InteropServices;',
	'public enum SW {',
	'    Hide            = 0,',
	'    Normal          = 1,',
	'    ShowMinimized   = 2,',
	'    Maximize        = 3,',
	'    ShowNoActivate  = 4,',
	'    Show            = 5,',
	'    Minimize        = 6,',
	'    ShowMinNoActive = 7,',
	'    ShowNA          = 8,',
	'    Restore         = 9,',
	'    ShowDefault     = 10,',
	'    ForceMinimize   = 11',
	'}',
	'public class Win32 {',
	'    [DllImport("user32.dll")]',
	'    public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);',
	'',
	'    [DllImport("user32.dll")]',
	'    public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);',
	'',
	'    [DllImport("user32.dll")]',
	'    public static extern bool IsWindowVisible(IntPtr hWnd);',
	'}'
) -Join "`n"
Add-Type $TypeDefinition -Language CSharp

[Diagnostics.Process]$Self = Get-Process -Id $Pid
[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Hide)

Add-Type -AssemblyName System.Windows.Forms, System.Drawing

[String]$DataPath     = "$Env:AppData\PSTrayUtil"
[String]$PreviousData = "$DataPath\Previous.xml"

:ParseInput Switch ($PSCmdlet.ParameterSetName) {
	'Name' {
		Try   {[String]$Path = (Get-Process -Name ($Name -Replace '\.exe$', '') -EA 1).Path}
		Catch {Throw "No process found by name '$Name'"}
		If (![IO.File]::Exists($Path)) {Throw "Error. Try using -Process, -PID or -Path instead."}
		Break
	}
	'PID' {
		Try   {[String]$Path = (Get-Process -Id $Id -EA 1).Path}
		Catch {Throw "No process found by PID '$Id'"}
		If (![IO.File]::Exists($Path)) {Throw "Error. Try using -Name, -Process or -Path instead."}
		Break
	}
	'InputObject' {
		If     ($InputObject -Is [Diagnostics.Process]) {[String]$Path = $InputObject.Path}
		ElseIf ($InputObject -As [Int])                 {[String]$Path = (Get-Process -Id ($InputObject -As [Int]) -EA 0).Path}
		ElseIf ($InputObject -Is [String]) {
			If     ($InputObject -Match '(?i)^([A-Z]:\\)|(\\\\)[\\/\S ]*\.exe$') {[String]$Path = $InputObject}
			ElseIf ($InputObject -Match '[^\\/](\.exe)?$')                       {[String]$Path = (Get-Process -Name ($InputObject -Replace '\.exe$', '') -EA 0).Path}
		}
		If (![IO.File]::Exists($Path)) {Throw "Error. Try using -Name, -PID or -Path instead."}
		Break
	}
	'Process' {
		[String]$Path = $Process.Path
		If (![IO.File]::Exists($Path)) {Throw "Error. Try using -Name, -PID or -Path instead."}
		Break
	}
	'Previous' {
		If (![IO.File]::Exists($PreviousData)) {Throw}
		[Hashtable]$Previous     = Import-Clixml -Path $PreviousData -EA 1
		[String]$Path            = $Previous.Path
		[Bool]$PreviouslyVisible = $Previous.Visible
		If ([IO.File]::Exists($Path)) {Break ParseInput}

		Remove-Item $PreviousData -Force -EA 0
		[Bool]$PreviouslyVisible = $True
		[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Show)

		:GetPath While ($True) {
			$Host.UI.RawUI.FlushInputBuffer()
			[String]$UserInput = Read-Host -Prompt 'Executable path or process name'
			If     ($UserInput -Match '(?i)^([A-Z]:\\)|(\\\\)[\\/\S ]*\.exe$') {[String]$Path = $UserInput}
			ElseIf ($UserInput -Match '[^\\/](\.exe)?$')                       {[String]$Path = (Get-Process -Name ($UserInput -Replace '\.exe$', '') -EA 0).Path}
			Else                                                               {[String]$Path = ''}
			If (![IO.File]::Exists($Path)) {Continue}
			[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Hide)
			Break ParseInput
		}
	}
	'Path'  {Break}
	Default {Throw "Invalid parameter set: $($PSCmdlet.ParameterSetName)"}
}

[Windows.Forms.NotifyIcon]$Icon = @{
	Icon = [Drawing.Icon]::ExtractAssociatedIcon($Self.Path)
	Text = 'PSTrayUtil'
}

If (![IO.Directory]::Exists($DataPath)) {[Void][IO.Directory]::CreateDirectory($DataPath)}

ForEach ($File in Get-ChildItem "$DataPath\*.txt" -File) {
	[Diagnostics.Process]$Instance = Get-Process -Id $File.BaseName -EA 0
	If (!$Instance) {[IO.File]::Delete($File.FullName)}
	ElseIf ([IO.File]::ReadAllText($File.FullName, [Text.UTF8Encoding]::New($False)) -Like $Path) {
		$Icon.Visible = $True
		$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Aborted:`nApplication is already managed.`nPID: $($File.BaseName)", [Windows.Forms.ToolTipIcon]::Error)
		$Icon.Visible = $False
		[Windows.Forms.Application]::Exit()
		Exit
	}
}

[Hashtable]$Previous = @{
	Path    = $Path
	Visible = $PreviouslyVisible
}
$Previous | Export-Clixml -Path $PreviousData -Force -EA 0

$Icon.Icon    = [Drawing.Icon]::ExtractAssociatedIcon($Path)
$Icon.Visible = $True

[Diagnostics.Process]$Target = Get-Process | Where-Object {$_.Path -eq $Path}
If (!$Target) {
	[Diagnostics.Process]$Target = Start-Process $Path -PassThru
	$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Started $($Target.ProcessName)", [Windows.Forms.ToolTipIcon]::Info)
	Start-Sleep -Seconds 1
}
Else {$Icon.ShowBalloonTip(1000, "$($Target.ProcessName) Tray", 'Ready', [Windows.Forms.ToolTipIcon]::Info)}
$Icon.Text = "$($Target.ProcessName) Tray"

[Threading.ParameterizedThreadStart]$WatcherProcedure = {
	Param (
		[Diagnostics.Process]$Proc,
		[Drawing.Icon]$Icon,
		[Management.Automation.PSReference]$ShouldStop
	)
	While (!$ShouldStop.Value) {
		$Proc.Refresh()
		If ($Proc.HasExited) {
			[Windows.Forms.Application]::Invoke({
				$Icon.Visible = $False
				[Windows.Forms.Application]::Exit()
				Return
			})
		}
		Start-Sleep -Milliseconds 250
	}
	Return
}
[Threading.Thread]$WatcherThread = [Threading.Thread]::New($WatcherProcedure)
$WatcherThread.IsBackground      = $True
[Bool]$StopWatcher               = $False

[IO.File]::WriteAllText("$DataPath\$Pid.txt", $Path, [Text.UTF8Encoding]::New($False))

$Target.Refresh()
[IntPtr]$TargetHandle = $Target.MainWindowHandle

If (!$PreviouslyVisible -And [Win32]::IsWindowVisible($TargetHandle)) {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Hide)}

$Target.PriorityClass = Switch ($Target.PriorityClass) {
	'RealTime'    {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
	'High'        {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
	'AboveNormal' {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
	'Normal'      {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
	Default       {$_; Break}
}

[Windows.Forms.ContextMenuStrip]$Menu         = [Windows.Forms.ContextMenuStrip]::New()
[Windows.Forms.ToolStripSeparator]$Separator1 = [Windows.Forms.ToolStripSeparator]::New()
[Windows.Forms.ToolStripSeparator]$Separator2 = [Windows.Forms.ToolStripSeparator]::New()
[Windows.Forms.ToolStripMenuItem]$ShowBtn     = [Windows.Forms.ToolStripMenuItem]::New("Show $($Target.ProcessName)")
[Windows.Forms.ToolStripMenuItem]$HideBtn     = [Windows.Forms.ToolStripMenuItem]::New("Hide $($Target.ProcessName)")
[Windows.Forms.ToolStripMenuItem]$RestartBtn  = [Windows.Forms.ToolStripMenuItem]::New("Restart $($Target.ProcessName)")
[Windows.Forms.ToolStripMenuItem]$QuitBtn     = [Windows.Forms.ToolStripMenuItem]::New("Quit $($Target.ProcessName)")
[Windows.Forms.ToolStripMenuItem]$ExitBtn     = [Windows.Forms.ToolStripMenuItem]::New("Exit PSTrayUtil")

$ShowBtn.add_Click({If (![Win32]::IsWindowVisible($TargetHandle)) {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Show)}})

$HideBtn.add_Click({If ([Win32]::IsWindowVisible($TargetHandle))  {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Hide)}})

$RestartBtn.add_Click({
	[Bool]$WasVisible = [Win32]::IsWindowVisible($TargetHandle)

	$SCRIPT:StopWatcher = $True
	While ($SCRIPT:WatcherThread.IsAlive) {Start-Sleep -Ms 100}

	$SCRIPT:Target.Refresh()
	If (!$SCRIPT:Target.HasExited) {Stop-Process $SCRIPT:Target -Force}

	[Diagnostics.Process]$SCRIPT:Target = Start-Process $Path -PassThru

	Start-Sleep -Seconds 1

	$SCRIPT:StopWatcher                     = $False
	[Threading.Thread]$SCRIPT:WatcherThread = [Threading.Thread]::New($WatcherProcedure)
	$SCRIPT:WatcherThread.IsBackground      = $True

	Try   {$SCRIPT:WatcherThread.Start(@($SCRIPT:Target, $Icon, [Ref]$StopWatcher))}
	Catch {$Icon.ShowBalloonTip(1000, 'PSTrayUtil', 'Failed to restart watcher thread.', [Windows.Forms.ToolTipIcon]::Warning)}

	$SCRIPT:Target.Refresh()
	[IntPtr]$SCRIPT:TargetHandle = $SCRIPT:Target.MainWindowHandle

	If (!$WasVisible -And [Win32]::IsWindowVisible($SCRIPT:TargetHandle))     {[Void][Win32]::ShowWindow($SCRIPT:TargetHandle, [SW]::Hide)}
	ElseIf ($WasVisible -And ![Win32]::IsWindowVisible($SCRIPT:TargetHandle)) {[Void][Win32]::ShowWindow($SCRIPT:TargetHandle, [SW]::Show)}

	$SCRIPT:Target.PriorityClass = Switch ($SCRIPT:Target.PriorityClass) {
		'RealTime'    {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
		'High'        {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
		'AboveNormal' {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
		'Normal'      {[Diagnostics.ProcessPriorityClass]::BelowNormal; Break}
		Default       {$_; Break}
	}

	$Icon.ShowBalloonTip(1000, $Target.ProcessName, 'Restarted', [System.Windows.Forms.ToolTipIcon]::Info)
})

$QuitBtn.add_Click({
	$SCRIPT:Previous['Visible'] = [Win32]::IsWindowVisible($TargetHandle)
	$SCRIPT:Previous | Export-Clixml -Path $PreviousData -Force -EA 0

	$SCRIPT:StopWatcher = $True
	While ($SCRIPT:WatcherThread.IsAlive) {Start-Sleep -Ms 100}

	$SCRIPT:Target.Refresh()
	If (!$SCRIPT:Target.HasExited) {Stop-Process $SCRIPT:Target -Force}

	$SCRIPT:Icon.Visible = $False

	[Windows.Forms.Application]::Exit()
})

$ExitBtn.add_Click({
	If (![Win32]::IsWindowVisible($TargetHandle)) {
		[Void][Win32]::ShowWindow($TargetHandle, [SW]::Show)
		$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Restored $($Target.ProcessName) window. Exiting.", [System.Windows.Forms.ToolTipIcon]::Info)
	}

	$SCRIPT:Icon.Visible = $False
	$SCRIPT:StopWatcher  = $True

	While ($SCRIPT:WatcherThread.IsAlive) {Start-Sleep -Ms 100}

	[Windows.Forms.Application]::Exit()
})

[Windows.Forms.ToolStripItem[]]$MenuItems = @(
	$ShowBtn,
	$HideBtn,
	$Separator1,
	$RestartBtn,
	$QuitBtn,
	$Separator2,
	$ExitBtn
)

[Void]$Menu.Items.AddRange($MenuItems)
$Icon.ContextMenuStrip = $Menu

Try   {$WatcherThread.Start(@($Target, $Icon, [Ref]$StopWatcher))}
Catch {$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Failed to start watcher thread.`n$($_.Exception.Message)", [Windows.Forms.ToolTipIcon]::Warning)}

[Windows.Forms.Application]::Run()
Exit
