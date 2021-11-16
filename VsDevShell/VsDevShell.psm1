
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
    $VsInstallPath = & $VsWherePath "-latest" "-property" "installationPath"
    $VsInstallPath
}

function Get-VsPSModulePath
{
    [CmdletBinding()]
    param(
    )
    
    $VsInstallPath = Get-VsInstallPath
    $VsToolsPath = Join-Path $VsInstallPath "Common7/Tools"
    $VsPSModulePath = Join-Path $VsToolsPath "Microsoft.VisualStudio.DevShell.dll"
    $VsPSModulePath
}

function Get-VsDevPath
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet('InstallPath','PSModulePath','vswhere')]
        [string] $PathType
    )
    
    switch ($PathType) {
        "InstallPath" { Get-VsInstallPath }
        "PSModulePath" { Get-VsPSModulePath }
        "vswhere" { Get-VsWherePath }
    }
}

function Import-VsDevShell
{
    [CmdletBinding()]
    param(
    )
    
    $VsPSModulePath = Get-VsDevPath 'PSModulePath'
    Import-Module $VsPSModulePath
}

function Enter-VsDevShell
{
    [CmdletBinding()]
    param(
        [switch] $NoLogo
    )

    if (-Not (Get-Module Microsoft.VisualStudio.DevShell)) {
        Import-VsDevShell
    }

    $VsInstallPath = Get-VsInstallPath

    $Params = @{
        VsInstallPath = $VsInstallPath;
        SkipExistingEnvironmentVariables = $true;
        SkipAutomaticLocation = $true;
    }
    
    if ($NoLogo) {
        Microsoft.VisualStudio.DevShell\Enter-VsDevShell @Params | Out-Null
    } else {
        Microsoft.VisualStudio.DevShell\Enter-VsDevShell @Params
    }
}
