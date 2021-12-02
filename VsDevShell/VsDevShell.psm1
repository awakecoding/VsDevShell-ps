
function Get-VsWherePath
{
    [CmdletBinding()]
    param(
    )

    $VsWhereCommand = Get-Command -Name vswhere -CommandType Application -ErrorAction SilentlyContinue

    if ($VsWhereCommand) {
        $VsWherePath = $VswhereCommand.Source
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

    $VsCmdOutput = & "${Env:COMSPEC}" "/c `"`"$VsDevCmdPath`" $VsCmdArgs && set"

    if ($LASTEXITCODE -ne 0) {
        throw "Failed to execute VsDevCmd.bat"
    }

    foreach ($VsCmdLine in $VsCmdOutput) {
        if ($VsCmdLine -Match '(.*)=(.*)') {
            $Name = $Matches[1]
            $Value = $Matches[2]
            [System.Environment]::SetEnvironmentVariable($Name, $Value)
        } else {
            Write-Host $VsCmdLine
        }
    }
}
