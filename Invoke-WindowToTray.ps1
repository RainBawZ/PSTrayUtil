Function New-WatcherInstance {
	[CmdletBinding()]

	Param (
		[Parameter(Mandatory)]
		[Diagnostics.Process]$Process,

		[Parameter(Mandatory)]
		[Ref]$Notify,

		[Parameter(Mandatory)]
		[Ref]$Flag
	)
	[Management.Automation.Runspaces.RunspacePool]$RunspacePool = [RunspaceFactory]::CreateRunspacePool(1, 1)
	$RunspacePool.Open()
	
	[PowerShell]$InitInstance   = [PowerShell]::Create()
	$InitInstance.RunspacePool  = $RunspacePool
	[PowerShell]$ScriptInstance = $InitInstance.AddScript({
		Param ([Diagnostics.Process]$Proc, [Ref]$Icon, [Ref]$Flag)
		Trap {
			$Flag.Value = 'Error'
			Write-Error "Watcher instance error: $($_.Exception.Message)"
			$Icon.Value.ShowBalloonTip(1000, 'PSTrayUtil', "Watcher error: $($_.Exception.Message)", [Windows.Forms.ToolTipIcon]::Error)
			Break
		}
		:Main While ($True) {
			Start-Sleep -Milliseconds 250
			Switch ($Flag.Value) {
				'Stop'      {$Flag.Value = 'Stopped'; Break Main}
				'Stopped'   {Break Main}
				'Suspend'   {$Flag.Value = 'Suspended'; Continue Main}
				'Suspended' {Continue Main}
				'Run'       {$Flag.Value = 'Running'; Break}
				'Running'   {Break}
				'Error'     {Continue Main}
				Default     {$Flag.Value = "`"$_`"?"; Break}
			}
			$Proc.Refresh()
			If ($Proc.HasExited) {
				$Icon.Value.Visible = $False
				[Windows.Forms.Application]::Exit()
				$Icon.Value.Dispose()
				$Flag.Value = 'Stopped'
				Break Main
			}
		}
		Exit
	})
	
	[PowerShell]$NewInstance = $ScriptInstance.AddArgument($Process).AddArgument([Ref]$Notify).AddArgument([Ref]$Flag)
	
	Try     {Return $NewInstance}
	Finally {
		Try {$RunspacePool.Close()}     Catch {}
		Try {$RunspacePool.Dispose()}   Catch {}
		Try {$InitInstance.Dispose()}   Catch {}
		Try {$ScriptInstance.Dispose()} Catch {}
	}
}

Function Stop-WatcherInstance {
	[CmdletBinding(DefaultParameterSetName = 'Default')]

	Param (
		[Parameter(Mandatory, ParameterSetName = 'Default')]
		[Parameter(Mandatory, ParameterSetName = 'Force')]
		[Ref]$Instance,

		[Parameter(Mandatory, ParameterSetName = 'Default')]
		[Parameter(Mandatory, ParameterSetName = 'Force')]
		[Ref]$Job,

		[Parameter(Mandatory, ParameterSetName = 'Default')]
		[Parameter(Mandatory, ParameterSetName = 'Force')]
		[Ref]$StopFlag,

		[Parameter(ParameterSetName = 'Default')]
		[UInt16]$Timeout = 5,

		[Parameter(Mandatory, ParameterSetName = 'Force')]
		[Switch]$Force
	)
	Try {
		$StopFlag.Value = $True
		Switch ($PSCmdlet.ParameterSetName) {
			'Default' {
				[DateTime]$StoppedAt   = [DateTime]::Now
				[TimeSpan]$StopTimeout = [TimeSpan]::FromSeconds($Timeout)
				While (!$Job.Value.IsCompleted -Or ([DateTime]::Now - $StoppedAt) -gt $StopTimeout) {Start-Sleep -Ms 250}
				If (!$Job.Value.IsCompleted) {
					Write-Warning 'Instance did not gracefully stop in time. Forcing stop.'
					$Instance.Value.Stop()
				}
				[Void]$Instance.Value.EndInvoke($Job.Value)
				Break
			}
			'Force' {
				$Instance.Value.Stop()
				[Void]$Instance.Value.EndInvoke($Job.Value)
				Break
			}
			Default {Throw "Invalid parameter set: $($PSCmdlet.ParameterSetName)"}
		}
		Return
	}
	Catch {
		[IO.File]::AppendAllText($ErrorLog, $UTF8, ($_.PSObject.Properties.Value | Out-String))
		Write-Error "Failed to stop instance: $_"
		Return
	}
	Finally {
		$StopFlag.Value = $False
		$Instance.Value.Dispose()
		$Instance.Value = $Null
		$Job.Value      = $Null
	}
}


Function Invoke-WindowToTray {
	[CmdletBinding(DefaultParameterSetName = 'Path')]

	Param (
		[Parameter(Mandatory, Position = 0, ParameterSetName = 'Path')]
		[ValidateScript({[IO.File]::Exists($_)})]
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

	[Diagnostics.Process]$Self = Get-Process -Id $Pid
	[Text.Encoding]$UTF8       = [Text.UTF8Encoding]::New($False)

	Trap {
		[IO.File]::AppendAllText($ErrorLog, $UTF8, ($_.PSObject.Properties.Value | Out-String))
		Write-Error $_
		Break
	}

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

	[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Hide)

	Add-Type -AssemblyName System.Windows.Forms, System.Drawing

	[Windows.Forms.NotifyIcon]$Icon = @{
		Icon    = [Drawing.Icon]::ExtractAssociatedIcon($Self.Path)
		Text    = 'PSTrayUtil'
		Visible = $False
	}

	[String]$DataPath     = "$Env:AppData\PSTrayUtil"
	[String]$ErrorLog     = "$DataPath\Error.log"
	[String]$PreviousData = "$DataPath\Previous.xml"

	:ParseInput Switch ($PSCmdlet.ParameterSetName) {
		'Name' {
			Try   {[String]$Path = (Get-Process -Name ($Name -Replace '\.exe$', '') -ErrorAction Stop).Path}
			Catch {Throw "No process found by name '$Name'"}
			If (![IO.File]::Exists($Path)) {Throw 'Error. Try using -Process, -PID or -Path instead.'}
			Break
		}
		'PID' {
			Try   {[String]$Path = (Get-Process -Id $Id -ErrorAction Stop).Path}
			Catch {Throw "No process found by PID '$Id'"}
			If (![IO.File]::Exists($Path)) {Throw 'Error. Try using -Name, -Process or -Path instead.'}
			Break
		}
		'InputObject' {
			If     ($InputObject -Is [Diagnostics.Process]) {[String]$Path = $InputObject.Path}
			ElseIf ($InputObject -As [UInt16])              {[String]$Path = (Get-Process -Id ($InputObject -As [Int]) -ErrorAction SilentlyContinue).Path}
			ElseIf ($InputObject -Is [String]) {
				If     ($InputObject -Match '(?i)^([A-Z]:\\)|(\\\\)[\\/\S ]*\.exe$') {[String]$Path = $InputObject}
				ElseIf ($InputObject -Match '[^\\/](\.exe)?$')                       {[String]$Path = (Get-Process -Name ($InputObject -Replace '\.exe$', '') -ErrorAction SilentlyContinue).Path}
			}
			If (![IO.File]::Exists($Path)) {Throw 'Error. Try using -Name, -PID or -Path instead.'}
			Break
		}
		'Process' {
			[String]$Path = $Process.Path
			If (![IO.File]::Exists($Path)) {Throw 'Error. Try using -Name, -PID or -Path instead.'}
			Break
		}
		'Previous' {
			If (![IO.File]::Exists($PreviousData)) {Throw}
			[Hashtable]$Previous     = Import-Clixml -Path $PreviousData -ErrorAction Stop
			[String]$Path            = $Previous.Path
			[Bool]$PreviouslyVisible = $Previous.Visible

			If ([IO.File]::Exists($Path)) {Break ParseInput}

			[IO.File]::Delete($PreviousData)
			[Bool]$PreviouslyVisible = $True
			[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Show)

			:GetPath While ($True) {
				$Host.UI.RawUI.FlushInputBuffer()
				[String]$UserInput = Read-Host -Prompt 'Executable path or process name'
				If     ($UserInput -Match '(?i)^([A-Z]:\\)|(\\\\)[\\/\S ]*\.exe$') {[String]$Path = $UserInput}
				ElseIf ($UserInput -Match '[^\\/](\.exe)?$')                       {[String]$Path = (Get-Process -Name ($UserInput -Replace '\.exe$', '') -ErrorAction SilentlyContinue).Path}
				Else                                                               {[String]$Path = ''}
				If (![IO.File]::Exists($Path)) {Continue}
				[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Hide)
				Break ParseInput
			}
		}
		'Path'  {Break}
		Default {Throw "Invalid parameter set: $($PSCmdlet.ParameterSetName)"}
	}

	If (![IO.Directory]::Exists($DataPath)) {[Void][IO.Directory]::CreateDirectory($DataPath)}

	ForEach ($File in Get-ChildItem -Path "$DataPath\*.txt" -File) {
		[Diagnostics.Process]$Instance = Get-Process -Id $File.BaseName -ErrorAction SilentlyContinue
		If     (!$Instance) {$File.Delete()}
		ElseIf ([IO.File]::ReadAllText($File, $UTF8) -Like $Path) {
			$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Aborted:`nApplication is already managed.`nPID: $($File.BaseName)", [Windows.Forms.ToolTipIcon]::Error)
			$Icon.Visible = $False
			[Windows.Forms.Application]::Exit()
			$Icon.Dispose()
			Return
		}
	}

	[Hashtable]$Previous = @{
		Path    = $Path
		Visible = $PreviouslyVisible
	}
	$Previous | Export-Clixml -Path $PreviousData -Force -ErrorAction SilentlyContinue

	[Diagnostics.Process]$Target = Get-Process | Where-Object {$_.Path -eq $Path}
	If (!$Target) {
		$Target     = Start-Process $Path -PassThru
		$TargetName = $Target.ProcessName
		$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Started $TargetName", [Windows.Forms.ToolTipIcon]::Info)
		Start-Sleep -Seconds 1
	}
	Else {$Icon.ShowBalloonTip(1000, "$TargetName Tray", 'Ready', [Windows.Forms.ToolTipIcon]::Info)}

	$Icon.Icon = [Drawing.Icon]::ExtractAssociatedIcon($Path)
	$Icon.Text = "$TargetName Tray"

	[IO.File]::WriteAllText("$DataPath\$Pid.txt", $Path, $UTF8)

	$Target.Refresh()
	[IntPtr]$TargetHandle = $Target.MainWindowHandle
	$Target.PriorityClass = [Diagnostics.ProcessPriorityClass]::BelowNormal

	If (!$PreviouslyVisible -And [Win32]::IsWindowVisible($TargetHandle)) {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Hide)}

	[Windows.Forms.ContextMenuStrip]$Menu         = [Windows.Forms.ContextMenuStrip]::New()
	[Windows.Forms.ToolStripSeparator]$Separator1 = [Windows.Forms.ToolStripSeparator]::New()
	[Windows.Forms.ToolStripSeparator]$Separator2 = [Windows.Forms.ToolStripSeparator]::New()
	[Windows.Forms.ToolStripSeparator]$Separator3 = [Windows.Forms.ToolStripSeparator]::New()
	[Windows.Forms.ToolStripSeparator]$Separator4 = [Windows.Forms.ToolStripSeparator]::New()
	[Windows.Forms.ToolStripMenuItem]$LogBtn      = [Windows.Forms.ToolStripMenuItem]::New('View error log')
	[Windows.Forms.ToolStripMenuItem]$PSTUBtn     = [Windows.Forms.ToolStripMenuItem]::New('Reveal PSTrayUtil Console')
	[Windows.Forms.ToolStripMenuItem]$ShowBtn     = [Windows.Forms.ToolStripMenuItem]::New("Show $($Target.ProcessName)")
	[Windows.Forms.ToolStripMenuItem]$HideBtn     = [Windows.Forms.ToolStripMenuItem]::New("Hide $($Target.ProcessName)")
	[Windows.Forms.ToolStripMenuItem]$RestartBtn  = [Windows.Forms.ToolStripMenuItem]::New("Restart $($Target.ProcessName)")
	[Windows.Forms.ToolStripMenuItem]$QuitBtn     = [Windows.Forms.ToolStripMenuItem]::New("Quit $($Target.ProcessName)")
	[Windows.Forms.ToolStripMenuItem]$ExitBtn     = [Windows.Forms.ToolStripMenuItem]::New('Exit PSTrayUtil')

	$LogBtn.add_Click({
		If ([IO.File]::Exists($ErrorLog)) {[Diagnostics.Process]::Start('notepad.exe', $ErrorLog)}
		Else {$Icon.ShowBalloonTip(1000, 'PSTrayUtil', 'No error log', [Windows.Forms.ToolTipIcon]::Info)}
	})

	$PSTUBtn.add_Click({
		If ([Win32]::IsWindowVisible($Self.MainWindowHandle)) {
			[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Hide)
			$Icon.ShowBalloonTip(1000, 'PSTrayUtil', 'Console hidden', [Windows.Forms.ToolTipIcon]::Info)
			$SCRIPT:PSTUBtn.Text = 'Reveal PSTrayUtil Console'
		}
		Else {
			[Void][Win32]::ShowWindow($Self.MainWindowHandle, [SW]::Show)
			$SCRIPT:PSTUBtn.Text = 'Hide PSTrayUtil Console'
		}
	})

	$ShowBtn.add_Click({If (![Win32]::IsWindowVisible($TargetHandle)) {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Show)}})

	$HideBtn.add_Click({If ([Win32]::IsWindowVisible($TargetHandle))  {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Hide)}})

	$RestartBtn.add_Click({
		[Bool]$WasVisible = [Win32]::IsWindowVisible($TargetHandle)

		Stop-WatcherInstance -Instance ([Ref]$SCRIPT:WatcherInstance) -Job ([Ref]$SCRIPT:WatcherJob) -Flag ([Ref]$SCRIPT:StopWatcher)

		$Target.Refresh()
		If (!$Target.HasExited) {$Target.CloseMainWindow()}
		$Target.WaitForExit(5000)
		
		[IO.File]::Delete("$DataPath\$($Target.Id).txt")

		[Diagnostics.Process]$Target = [Diagnostics.Process].Start($Path)
		[String]$TargetName          = $SCRIPT:Target.ProcessName

		Start-Sleep -Seconds 1

		$Target.Refresh()
		$TargetHandle = $Target.MainWindowHandle

		If     (!$WasVisible -And [Win32]::IsWindowVisible($TargetHandle)) {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Hide)}
		ElseIf ($WasVisible -And ![Win32]::IsWindowVisible($TargetHandle)) {[Void][Win32]::ShowWindow($TargetHandle, [SW]::Show)}

		$Target.PriorityClass = [Diagnostics.ProcessPriorityClass]::BelowNormal

		Try {
			#[PowerShell]$NewWatcherInstance  = New-WatcherInstance -Process $SCRIPT:Target -Notify ([Ref]$SCRIPT:Icon) -Flag ([Ref]$SCRIPT:StopWatcher)
			#$SCRIPT:WatcherInstance          = $NewWatcherInstance
			#[IAsyncResult]$SCRIPT:WatcherJob = $SCRIPT:WatcherInstance.BeginInvoke()

			$Icon.ShowBalloonTip(1000, $TargetName, 'Restarted', [Windows.Forms.ToolTipIcon]::Info)
		}
		Catch {
			[IO.File]::AppendAllBytes($ErrorLog, $UTF8.GetBytes(($_.PSObject.Properties.Value | Out-String)))
			$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Failed to restart watcher thread.`n$($_.Exception.Message)", [Windows.Forms.ToolTipIcon]::Warning)
		}
	})

	$QuitBtn.add_Click({
		$SCRIPT:Previous['Visible'] = [Win32]::IsWindowVisible($TargetHandle)
		$SCRIPT:Previous | Export-Clixml -Path $PreviousData -Force -ErrorAction SilentlyContinue

		Stop-WatcherInstance -Instance ([Ref]$SCRIPT:WatcherInstance) -Job ([Ref]$SCRIPT:WatcherJob) -Flag ([Ref]$SCRIPT:StopWatcher)

		$SCRIPT:Icon.Visible = $False

		[Windows.Forms.Application]::Exit()
	})

	$ExitBtn.add_Click({
		If (![Win32]::IsWindowVisible($TargetHandle)) {
			[Void][Win32]::ShowWindow($TargetHandle, [SW]::Show)
			$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Restoring $($Target.ProcessName) window and exiting.", [System.Windows.Forms.ToolTipIcon]::Info)
		}

		Stop-WatcherInstance -Instance ([Ref]$SCRIPT:WatcherInstance) -Job ([Ref]$SCRIPT:WatcherJob) -Flag ([Ref]$SCRIPT:StopWatcher)

		$SCRIPT:Icon.Visible = $False

		[Windows.Forms.Application]::Exit()
	})

	[Windows.Forms.ToolStripItem[]]$MenuItems = @(
		$PSTUBtn,
		$Separator1,
		$ShowBtn,
		$HideBtn,
		$Separator2,
		$RestartBtn,
		$QuitBtn,
		$Separator3,
		$LogBtn,
		$Separator4,
		$ExitBtn
	)

	[Void]$Menu.Items.AddRange($MenuItems)
	$Icon.ContextMenuStrip = $Menu

	[Bool]$StopWatcher = $False
	Try {
		[PowerShell]$WatcherInstance = New-WatcherInstance -Process $Target -Notify ([Ref]$Icon) -Flag ([Ref]$StopWatcher)
		[IAsyncResult]$WatcherJob    = $WatcherInstance.BeginInvoke()
	}
	Catch {
		[IO.File]::AppendAllBytes($ErrorLog, $UTF8.GetBytes(($_.PSObject.Properties.Value | Out-String)))
		$Icon.ShowBalloonTip(1000, 'PSTrayUtil', "Failed to start watcher thread.`n$($_.Exception.Message)", [Windows.Forms.ToolTipIcon]::Warning)
	}

	[Windows.Forms.Application]::Run()

	Stop-WatcherInstance -Instance ([Ref]$WatcherInstance) -Job ([Ref]$WatcherJob) -Flag ([Ref]$StopWatcher)

	$Icon.Dispose()
	$Menu.Dispose()

	Return
}
