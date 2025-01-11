#!/usr/bin/env pwsh
##############################################################################################################

Function Show-Usage {
    @"
vagrant  = 'it-gro/win10-ltsc-eval'
download = 'https://microsoft.com/en-us/evalcenter'
package  = 'https://learn.microsoft.com/en-us/mem/configmgr/develop/apps/how-to-create-the-windows-installer-file-msi'
shell    = 'https://learn.microsoft.com/en-us/powershell'

Usage: pwsh -File $PSCommandPath [OPTIONS]
Options:
    build
    lint
"@ | Out-Host
}

Function Build-Project {
    'LOAD MAKE.JSON' | Out-Log
    New-Variable -Option Constant -Name VAR -Value (
        Get-Content -Path $PSCommandPath.Replace('ps1', 'json') | ConvertFrom-Json
    )
    'CHECK PATH WITH .LPI' | Out-Log
    If (! (Test-Path -Path $Var.app)) {
        "$Var.app did not find!" | Out-Log
        Exit 1
    }
    'UPDATE GIT SUBMODULE' | Out-Log
    If (Test-Path -Path '.gitmodules') {
        & git submodule update --init --recursive --force --remote | Out-Host
        "[$LastExitCode] git submodule update" | Out-Log
    }
    'INSTALL LAZBUILD' | Out-Log
    @(
        @{
            Cmd = 'lazbuild'
            Url = 'https://fossies.org/windows/misc/lazarus-3.6-fpc-3.2.2-win64.exe'
            Path = "C:\Lazarus"
        }
    ) | Where-Object {
        ! (Test-Path -Path $_.Path)
    } | ForEach-Object {
        $_.Url | Request-File | Install-Program
        $Env:PATH+=";$_.Path"
        Return (Get-Command $_.Cmd).Source
    } | Out-Host
    'LOAD PACKAGES FROM OPM' | Out-Log
    $VAR.Pkg | ForEach-Object {
        @{
            Name = $_
            Uri = "https://packages.lazarus-ide.org/$_.zip"
            Path = "$Env:APPDATA\.lazarus\onlinepackagemanager\packages\$_"
            OutFile = (New-TemporaryFile).FullName
        }
    } | Where-Object {
        ! (Test-Path -Path $_.Path) &&
        ! (& lazbuild --verbose-pkgsearch $_.Name ) &&
        ! (& lazbuild --add-package $_.Name)
    } | ForEach-Object -Parallel {
        Invoke-WebRequest -OutFile $_.OutFile -Uri $_.Uri
        New-Item -Type Directory -Path $_.Path | Out-Null
        Expand-Archive -Path $_.OutFile -DestinationPath $_.Path
        Remove-Item $_.OutFile
        (Get-ChildItem -Filter '*.lpk' -Recurse -File –Path $_.Path).FullName |
            ForEach-Object {
                & lazbuild --add-package-link $_ | Out-Null
                "[$LastExitCode] add package link $_" | Out-Log
            }
    }
    If (Test-Path -Path $VAR.lib) {
        'LOAD OTHER .LPK' | Out-Log
        (Get-ChildItem -Filter '*.lpk' -Recurse -File –Path $VAR.lib).FullName |
            Where-Object {
                $_ -notmatch '(cocoa|x11|_template)'
            } | ForEach-Object {
                & lazbuild --add-package-link $_ | Out-Null
               "[$LastExitCode] add package link $_" | Out-Log
            }
    }
    Exit $(
        If (Test-Path -Path $Var.tst) {
            'RUN UNITTEST' | Out-Log
            $Output = (
                & lazbuild --build-all --recursive --no-write-project $VAR.tst |
                    Where-Object {
                        $_ -match 'Linking'
                    } | ForEach-Object {
                        $_.Split(' ')[2].Replace('bin', 'bin\.')
                    }
            )
            & $Output --all --format=plain --progress | Out-Log
            If ($LastExitCode -eq 0) { 0 } Else { 1 }
        }
    ) + ('BUILD APPLICATIONS' | Out-Log) + (
        (Get-ChildItem -Filter '*.lpi' -Recurse -File –Path $Var.app).FullName |
            ForEach-Object {
                $Output = (& lazbuild --build-all --recursive --no-write-project $_)
                $Result = @("[$LastExitCode] $_")
                $exitCode = If ($LastExitCode -eq 0) {
                    $Result += $Output | Where-Object { $_ -match 'Linking' }
                    0
                } Else {
                    $Result += $Output | Where-Object { $_ -match '(Error|Fatal):' }
                    1
                }
                $Result | Out-Log
                Return $exitCode
            } | Measure-Object -Sum
    ).Sum
}

Filter Out-Log {
    $(
        If (Test-Path -Path Variable:LastExitCode) {
            "$([char]27)[33m$(Get-Date -uformat '%y-%m-%d_%T')`t{0}$([char]27)[0m" -f $_
        } ElseIf ($LastExitCode -eq 0) {
            "$([char]27)[32m$(Get-Date -uformat '%y-%m-%d_%T')`t{0}$([char]27)[0m" -f $_
        } Else {
            "$([char]27)[31m$(Get-Date -uformat '%y-%m-%d_%T')`t{0}$([char]27)[0m" -f $_
        }
    ) | Out-Host
}

Filter Request-File {
    New-Variable -Option Constant -Name VAR -Value @{
        Uri = $_
        OutFile = (Split-Path -Path $_ -Leaf).Split('?')[0]
    }
    Invoke-WebRequest @VAR
    Return $VAR.OutFile
}

Filter Install-Program {
    If ((Split-Path -Path $_ -Leaf).Split('.')[-1] -eq 'msi') {
        & msiexec /passive /package $_ | Out-Null
    } Else {
        & ".\$_" /SP- /VERYSILENT /SUPPRESSMSGBOXES /NORESTART  | Out-Null
    }
    Remove-Item $_
}

Function Request-URL([Switch] $Post) {
    $VAR = If ($Post) {
        @{
            Method = 'POST'
            Headers = @{
                ContentType = 'application/json'
            }
            Uri = 'https://postman-echo.com/post'
            Body = @{
                One = '1'
            } | ConvertTo-Json
        }
    } Else {
        @{
            Uri = 'https://postman-echo.com/get'
        }
    }
    Return (Invoke-WebRequest @VAR | ConvertFrom-Json).Headers
}

Function Switch-Action {
    $ErrorActionPreference = 'stop'
    Set-PSDebug -Strict #-Trace 1
    Invoke-ScriptAnalyzer -EnableExit -Path $PSCommandPath
    If ($args.count -gt 0) {
        If ($args[0] -eq 'build') {
            Build-Project
        } Else {
            Invoke-ScriptAnalyzer -EnableExit -Recurse -Path '.'
            (Get-ChildItem -Filter '*.ps1' -Recurse -Path '.').FullName |
                ForEach-Object {
                    Invoke-Formatter -ScriptDefinition $(
                        Get-Content -Path $_ | Out-String
                    ) | Set-Content -Path $_
                }
        }
    } Else {
        Show-Usage
    }
}

##############################################################################################################
Switch-Action @args
