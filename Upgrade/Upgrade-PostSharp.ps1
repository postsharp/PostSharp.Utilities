<#
.Synopsis
    Batch PostSharp upgrade.

.Description
    This script upgrades PostSharp to version specified by -postSharpVersion parameter in all
    projects (csproj, vbproj) recursively found in directory specified by -path parameter.
    Supported are projects referencing PostSharp 2.x, 3.x and 4.x. Projects without PostSharp are
    not processed.

    Script requires at least PowerShell 3 and Microsoft.Build v4.0. Script can run on machine
    without Visual Studio.

    Limitations:
        - Only .NET Framework targets are supported.
        - Projects referencing PostSharp toolkits are skipped.
        - Script must be run from directory with nuget.exe.
#>



param(
    [string]$path = $null,
    [string]$postSharpVersion = $null,
    [string]$outputProjectFileNameSuffix = '',
    [bool]$backup = $false
)

if (!(Test-Path .\NuGet.exe))
{
    Write-Error "NuGet.exe file not found. Run the script from directory that contains Nuget.exe"
    return
}

Add-Type -AssemblyName 'Microsoft.Build, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a'

function Get-RelativePath($basePath, $targetPath)
{
    $originalPwd = $PWD
    cd $basePath

    $relativePath = Resolve-Path $targetPath -Relative

    cd $originalPwd

    return $relativePath
}

function Get-PostSharpReference($csproj)
{
    return @($csproj.Xml.ItemGroups.Children | Where-Object { $_.ItemType -like 'Reference' -and $_.Include.ToLowerInvariant().StartsWith("postsharp") })
}

function Get-RepositoryPath($path)
{
    $repositoryPath = .\NuGet.exe config repositorypath
    if ($repositoryPath -like 'WARNING*')
    {
        Write-Warning 'repositorypath nuget setting is not set, using solution path as root for repository path'
        $repositoryPath = Join-Path $path 'packages'
    }

    Write-Host "Using $repositorypath as repositorypath"

    return $repositoryPath
}

function Upgrade-Project
{
    param(
        [string]$projectFullName,
        [string]$packagePath,
        [string]$outputProjectFullName = $null,
        [bool]$backup = $false,
        [string]$version
    )

    Write-Host "Processing $projectFullName"

    # Overwrite original project file if no output file name specified
    if ($outputProjectFullName -eq $null)
    {
        $outputProjectFullName = $projectFullName
    }

    # load the project
    try
    {
        $csproj = (New-Object Microsoft.Build.Evaluation.ProjectCollection).LoadProject($projectFullName)
    }
    catch [Exception]
    {
        Write-Error $_.Exception.Message
        Write-Warning 'Skipping project'
        return
    }

    # check project target
    $targetFrameworkIdentifier = $csproj.GetProperty('TargetFrameworkIdentifier')
    if ($targetFrameworkIdentifier -and $targetFrameworkIdentifier.EvaluatedValue -ne '.NETFramework')
    {
        Write-Warning "Project doesn't target .NET Framework. Skipping the project."
        return
    }

    # check if there are some references to PostSharp
    $postSharpReferences = Get-PostSharpReference $csproj

    if (!$postSharpReferences -or $postSharpReferences.length -eq 0)
    {
        Write-Warning "No PostSharp reference. Skipping the project"
        return
    }

    $toolkitReference = $postSharpReferences | Where-Object { $_.Include -like 'postsharp.toolkit*' }
    if ($toolkitReference -and $toolkitReference.lenght -ne 0)
    {
        Write-Warning "Project contains unsupported toolkit reference(s):"
        $toolkitReference | ForEach-Object { Write-Warning $_.Include }
        Write-Warning "Skipping the project."
        Write-Host ''
        return
    }

    $referenceGroup = $postSharpReferences[0].Parent

    $projectPath = Split-Path $projectFullName
    $relativePackagePath = Get-RelativePath $projectPath $packagePath
    $relativePostSharpTargetsPath = Join-Path $relativePackagePath 'tools\PostSharp.targets'
    $relativeAssemblyPath = Join-Path $relativePackagePath 'lib\net20\PostSharp.dll'
    $absoluteAssemblyPath = Join-Path $packagePath 'lib\net20\PostSharp.dll'

    # Remove elements from previous installations or versions.
    $nodesToRemove = @()
    $nodesToRemove += $csproj.Xml.Properties | Where-Object {$_.Name.ToLowerInvariant() -eq "dontimportpostsharp" }
    $nodesToRemove += $csproj.Xml.Imports | Where-Object {$_.Project.ToLowerInvariant().EndsWith("postsharp.targets") } 
    $nodesToRemove += $csproj.Xml.Targets | Where-Object {$_.Name.ToLowerInvariant() -eq "ensurepostsharpimported" }
    $nodesToRemove += $postSharpReferences

    $postSharpReferences | ForEach-Object { Write-Host "Removing reference" $_.Include }
    $nodesToRemove | ForEach-Object { $_.Parent.RemoveChild($_) | out-null }

    # Set property DontImportPostSharp to prevent locally-installed previous versions of PostSharp to interfere.
    $csproj.Xml.AddProperty( "DontImportPostSharp", "True" ) | Out-Null

    # Add import to PostSharp.targets
    $importGroup = $csproj.Xml.AddImportGroup() # make sure that PostSharp.targets is imported as a last
    $import = $importGroup.AddImport($relativePostSharpTargetsPath)
    $import.set_Condition( "Exists('$relativePostSharpTargetsPath')" ) | Out-Null

     # Add a target to fail the build when our targets are not imported
    $target = $csproj.Xml.AddTarget("EnsurePostSharpImported")
    $target.BeforeTargets = "BeforeBuild"
    $target.Condition = "'`$(PostSharp30Imported)' == ''"

    # if the targets don't exist at the time the target runs, package restore didn't run
    $errorTask = $target.AddTask("Error")
    $errorTask.Condition = "!Exists('$relativePostSharpTargetsPath')"
    $errorTask.SetParameter("Text", "This project references NuGet package(s) that are missing on this computer. Enable NuGet Package Restore to download them.  For more information, see http://www.postsharp.net/links/nuget-restore.");

    # if the targets exist at the time the target runs, package restore ran but the build didn't import the targets.
    $errorTask = $target.AddTask("Error")
    $errorTask.Condition = "Exists('$relativePostSharpTargetsPath')"
    $errorTask.SetParameter("Text", "The build restored NuGet packages. Build the project again to include these packages in the build. For more information, see http://www.postsharp.net/links/nuget-restore.");

    # get full name of PostSharp assemlby
    $assembly = [System.Reflection.AssemblyName]::GetAssemblyName($absoluteAssemblyPath)
    $include = $assembly.FullName
    if ($assembly.ProcessorArchitecture)
    {
        $include += ", processorArchitecture=" + $assembly.ProcessorArchitecture
    }

    $reference = $csproj.Xml.CreateItemElement("Reference")
    $reference.Include = $include
    $referenceGroup.AppendChild($reference)

    $reference.AppendChild($csproj.Xml.CreateMetadataElement("Private", "True"))
    $reference.AppendChild($csproj.Xml.CreateMetadataElement("HintPath", $relativeAssemblyPath))

    # backup original file
    if ($backup)
    {
        Copy-Item $projectFullName ($projectFullName + ".bak")
    }

    # save the project file
    $csproj.Save($outputProjectFullName)

    Upgrade-Packages -projectFullName $projectFullName -version $version -outpuSuffix $outputProjectFileNameSuffix
}

function Upgrade-Packages
{
    param(
        [string]$projectFullName,
        [string]$version,
        [string]$outpuSuffix = ''
    )

    $projectPath = Split-Path $projectFullName
    $packageFullName = Join-Path $projectPath 'packages.config'
    $outputFullName = $packageFullName + $outputSuffix

    if (Test-Path $packageFullName)
    {
        $xml = [xml](Get-Content $packageFullName)

        # backup original file
        if ($backup)
        {
            Copy-Item $packageFullName ($packageFullName + ".bak")
        }
    }
    else
    {
        Write-Host "Creating packages.config"
        $xml = [xml]"<packages></packages>"
    }

    

    $packageElement = $xml.packages.ChildNodes | Where-Object { $_.Name -like 'package' -and $_.id -like 'PostSharp' }[0]
    if (!$packageElement)
    {
        Write-Host "Installing NuGet package version $version"
        $packageElement = $xml.CreateElement("package")
        $packageElement.SetAttribute("id", "PostSharp")
        $packageElement.SetAttribute("version", $version)
        $packageElement.SetAttribute("targetFramework", "net20")
        $xml.DocumentElement.AppendChild($packageElement) | Out-Null
    }
    else
    {
        Write-Host "Updating current NuGet package version from" $packageElement.version "to $version"
        $packageElement.SetAttribute("version", $version)
    }

    

    

   

    $xml.Save($outputFullName)
}

function Upgrade-Directory
{
    param(
        [string]$rootPath = $null,
        [string]$postSharpVersion = $null,
        [string]$outputProjectFileNameSuffix = '',
        [bool]$backup = $true
    )
    try
    {
        $repositoryPath = Get-RepositoryPath $rootPath
        $nugetPackage = 'PostSharp.' + $postSharpVersion

        # Install nuget package    
        $nugetOutput = .\NuGet.exe install 'PostSharp' -Version $postSharpVersion -OutputDirectory $repositoryPath
        if (!($nugetOutput -like 'Successfully*') -and !($nugetOutput -like '*already installed.'))
        {
            Write-Warning 'PostSharp NuGet package not installed successfully. Terminating script.'
            return
        }
        Write-Host $nugetOutput
    
        # Update projects
        $postSharpNugetPath = Join-Path $repositoryPath $nugetPackage

        Write-Host ''
        Write-Host 'Updating projects'

        Get-ChildItem $rootPath -Recurse |
            Where-Object { ($_.Name -like '*.csproj') -or ($_.Name -like '*.vbproj') } |
            ForEach-Object {
                Upgrade-Project -projectFullName $_.FullName -packagePath $postSharpNugetPath -outputProjectFullName ($_.FullName + $outputProjectFileNameSuffix) -backup $backup -version $postSharpVersion
                Write-Host ''
            }
    }
    catch [Exception]
    {
        Write-Warning 'Unhandled exception thrown. Terminating batch upgrade.'
        throw
    }
}

if ($path -and $postSharpVersion)
{
    Upgrade-Directory -rootPath $path -postSharpVersion $postSharpVersion -outputProjectFileNameSuffix $outputProjectFileNameSuffix -backup $backup
}
else
{
    Write-Host '-path and -postSharpVersion are mandatory parameters. Please, specify them in order to start the batch upgrade.'
}
