
function Get-VsWherePath
{
    [CmdletBinding()]
    param(
    )

    $VsWhereCommand = Get-Command -Name vswhere -CommandType Application -ErrorAction SilentlyContinue

    if ($VsWhereCommand) {
        $VsWherePath = $VswhereCommand[0].Source
    } else {
        $VsWherePath = Join-Path ${Env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    }

    $VsWherePath
}

function Get-VsInstallPath
{
    [CmdletBinding()]
    param(
    )
    
    $VsWherePath = Get-VsWherePath

    if (-Not (Test-Path -Path $VsWherePath -PathType Leaf)) {
        throw [System.IO.FileNotFoundException] "vswhere.exe not found."
    }

    $VsInstallPath = & $VsWherePath "-latest" "-property" "installationPath"
    $VsInstallPath
}

function Get-VsDevCmdPath
{
    [CmdletBinding()]
    param(
        [string] $VsInstallPath
    )
    
    if ([string]::IsNullOrEmpty($VsInstallPath)) {
        $VsInstallPath = Get-VsInstallPath
    }

    $VsToolsPath = Join-Path $VsInstallPath "Common7/Tools"
    $VsDevCmdPath = Join-Path $VsToolsPath "VsDevCmd.bat"
    $VsDevCmdPath
}

function Get-VsDevEnv
{
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
        [ValidateSet('x86','x64','arm','arm64')]
        [string] $Arch = "x64",
        [ValidateSet('x86','x64')]
        [string] $HostArch = "x64",
        [ValidateSet('Desktop','UWP')]
        [string] $AppPlatform = "Desktop",
        [string] $WinSdk,
        [switch] $NoExt,
        [switch] $NoLogo,
        [string] $VsInstallPath
    )

    if ([string]::IsNullOrEmpty($VsInstallPath)) {
        $VsInstallPath = Get-VsInstallPath
    }

    if (-Not (Test-Path -Path $VsInstallPath -PathType Container)) {
        throw [System.IO.FileNotFoundException] "$VsInstallPath not found."
    }

    $VsDevCmdPath = Get-VsDevCmdPath -VsInstallPath $VsInstallPath

    if (-Not (Test-Path -Path $VsDevCmdPath -PathType Leaf)) {
        throw [System.IO.FileNotFoundException] "$VsDevCmdPath not found."
    }

    $Arch = $Arch.ToLower()
    $HostArch = $HostArch.ToLower()

    $VsCmdArgs = "-arch=$Arch"
    $VsCmdArgs += " -host_arch=$HostArch"

    if (-Not [string]::IsNullOrEmpty($WinSdk)) {
        $VsCmdArgs += " -winsdk=$WinSdk"
    }

    if ($NoExt) {
        $VsCmdArgs += " -no_ext"
    }

    if ($NoLogo) {
        $VsCmdArgs += " -no_logo"
    }

    $Env:VSCMD_SKIP_SENDTELEMETRY = "1"
    $Env:VSCMD_BANNER_SHELL_NAME_ALT = "$Arch Developer Shell"

    $processStartInfo = New-Object System.Diagnostics.ProcessStartInfo
    $processStartInfo.FileName = "${Env:COMSPEC}"
    $processStartInfo.Arguments = "/c `"`"$VsDevCmdPath`" $VsCmdArgs && set`""
    $processStartInfo.WorkingDirectory = Split-Path $VsDevCmdPath
    $processStartInfo.RedirectStandardOutput = $true
    $processStartInfo.UseShellExecute = $false
    $processStartInfo.CreateNoWindow = $true

    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $processStartInfo
    $process.Start() | Out-Null
    $VsCmdOutput = $process.StandardOutput.ReadToEnd()
    $process.WaitForExit()

    $VsCmdOutput = $VsCmdOutput -split "`r`n"

    if ($process.ExitCode -ne 0) {
        throw "Failed to execute VsDevCmd.bat"
    }

    $PreEnv = [ordered]@{}
    (Get-ChildItem env:) | ForEach-Object {
        $PreEnv.Add($_.Name, $_.Value)
    }

    $VsDevEnv = [ordered]@{}
    foreach ($VsCmdLine in $VsCmdOutput) {
        if ($VsCmdLine.Contains('=')) {
            $Name, $Value = $VsCmdLine -split '=', 2
            if ($PreEnv[$Name] -ne $Value) {
                $VsDevEnv.Add($Name, $Value)
            }
        }
    }

    $VsDevEnv
}

function Enter-VsDevShell
{
    [CmdletBinding()]
    param(
        [Parameter(Position=0)]
        [ValidateSet('x86','x64','arm','arm64')]
        [string] $Arch = "x64",
        [ValidateSet('x86','x64')]
        [string] $HostArch = "x64",
        [ValidateSet('Desktop','UWP')]
        [string] $AppPlatform = "Desktop",
        [string] $WinSdk,
        [switch] $NoExt,
        [switch] $NoLogo,
        [string] $VsInstallPath
    )

    $VsDevEnv = Get-VsDevEnv -Arch:$Arch -HostArch:$HostArch `
        -AppPlatform:$AppPlatform -WinSdk:$WinSdk `
        -NoExt:$NoExt -NoLogo:$NoLogo `
        -VsInstallPath:$VsInstallPath

    $VsDevEnv.GetEnumerator() | ForEach-Object {
        [System.Environment]::SetEnvironmentVariable($_.Key, $_.Value)
    }
}
