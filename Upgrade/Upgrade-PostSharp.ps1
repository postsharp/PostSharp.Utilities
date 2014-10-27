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
        - Projects referencing PostSharp Toolkits 2.1 and PostSharp Pattern Libraries are skipped.
        - NuGet.exe should be in present in the same directory as the script.
#>



param(
    [string]$path = $null,
    [string]$postSharpVersion = $null,
    [string]$outputProjectFileNameSuffix = '',
    [bool]$backup = $false,
    [string] $checkout = 'tf checkout "{0}"'
)

# Check PowerShell version
if ( $PSVersionTable.PSVersion.Major -lt 3 )
{
    Write-Error "This script requires PowerShell 3.0"
    return
}

if ( $PSVersionTable.CLRVersion.Major -lt 4 )
{
    Write-Error "This script requires CLR v4.0"
    return
}



# Check for presence of NuGet.exe

$nugetExe = Join-Path $(Split-Path -parent $MyInvocation.MyCommand.Definition) "nuget.exe"

if (!(Test-Path $nugetExe))
{
    Write-Error "NuGet.exe file not found. Run the script from directory that contains Nuget.exe"
    return
}

Add-Type -AssemblyName 'Microsoft.Build, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a' -ErrorAction Stop


function Checkout-File($file)
{
    $status = Get-ChildItem $file
    if ( $status.IsReadOnly )
    {
        Write-Warning "Checking out file $file"
        $checkoutCommandLine = $checkout -f $file
        & cmd /c $checkoutCommandLine
    }
}

function Get-RelativePath($basePath, $targetPath)
{
    $originalPwd = $PWD
    cd $basePath

    $relativePath = Resolve-Path $targetPath -Relative

    cd $originalPwd

    return $relativePath
}

function Has-PostSharpReference([Microsoft.Build.Evaluation.Project] $csproj)
{
   $projectInstance = $csproj.CreateProjectInstance()

    if ( !$projectInstance.Build("ResolveAssemblyReferences", $null) )
    {
        throw "Cannot resolve assembly references"
    }

    return $projectInstance.Items | ` 
        Where-Object   {  ( $_.ItemType -eq "ReferencePath" -or $_.ItemType -eq "ReferenceDependencyPaths") -and  `
                          ( $_.Metadata["FusionName"].EvaluatedValue -match "PostSharp, Version=.*, Culture=neutral, PublicKeyToken=b13fd38b8f9c99d7" )   }

}

function Get-DirectPostSharpReferences([Microsoft.Build.Evaluation.Project] $csproj)
{
 
    return @($csproj.Xml.ItemGroups.Children | Where-Object { $_.ItemType -like 'Reference' -and $_.Include.ToLowerInvariant().StartsWith("postsharp") })
}

function Get-RepositoryPath($path)
{
    Push-Location $path

    $repositoryPath = & $nugetExe config repositorypath

    if ($repositoryPath -like 'WARNING*')
    {
        Write-Warning 'repositorypath nuget setting is not set, using solution path as root for repository path'
        $repositoryPath = Join-Path $path 'packages'
    }
    else
    {
        $repositoryPath = Join-Path $path $repositoryPath
    }

    Write-Host "Using $repositorypath as repositorypath"

    Pop-Location

    return $repositoryPath
}

function Add-Reference
{
    param(
        [Microsoft.Build.Evaluation.Project] $csproj,
        [string] $packagePath,
        [string] $libraryPath
     )


    $projectDirectory = Split-Path $csproj.FullPath
    $relativePackagePath = Get-RelativePath $projectDirectory $packagePath
    $relativeAssemblyPath = Join-Path $relativePackagePath $libraryPath
    $absoluteAssemblyPath = Join-Path $packagePath $libraryPath

    if ( -not (Test-Path $absoluteAssemblyPath ) )
    {
        Write-Error "File $absoluteAssemblyPath does not exist."
        Quit
    }

    

    $assembly = [System.Reflection.AssemblyName]::GetAssemblyName($absoluteAssemblyPath)
    $include = $assembly.FullName
    if ($assembly.ProcessorArchitecture)
    {
        $include += ", processorArchitecture=" + $assembly.ProcessorArchitecture
    }

    # Add reference to project file.

    $reference = $csproj.Xml.CreateItemElement("Reference")
    $reference.Include = $include
    $referenceGroup.AppendChild($reference)

    $reference.AppendChild($csproj.Xml.CreateMetadataElement("Private", "True"))
    $reference.AppendChild($csproj.Xml.CreateMetadataElement("HintPath", $relativeAssemblyPath))



}

function Install-Package
{
    param(
        [string] $repositoryPath,
        [string] $packageName,
        [string] $packageVersion,
        [xml] $packagesConfig
        )

    $nugetPackage = $packageName + '.' + $postSharpVersion
    $postSharpNugetPath = Join-Path $repositoryPath $nugetPackage

    # Download the package if necessary.
    if ( -not (Test-Path $postSharpNugetPath ) )
    {
        Write-Host "Downloading $packageName $packageVersion ..."
    
        # Install nuget package    
        $nugetOutput = & $nugetExe install $packageName -Version $postSharpVersion -OutputDirectory $repositoryPath
        if (!($nugetOutput -like 'Successfully*') -and !($nugetOutput -like '*already installed.'))
        {
            Write-Error $nugetOutput
            Write-Warning "$nugetPackage NuGet package not installed successfully. Terminating script."
            return
        }
        Write-Host $nugetOutput
    
    }

    # Add to packages.config.
    $packageElement = $packagesConfig.packages.ChildNodes | Where-Object { $_.Name -eq 'package' -and $_.id -eq $packageName }[0]
    if (!$packageElement)
    {
        Write-Host "Installing NuGet package version $version"
        $packageElement = $xml.CreateElement("package")
        $packageElement.SetAttribute("id", $packageName)
        $packageElement.SetAttribute("version", $packageVersion)
        $packageElement.SetAttribute("targetFramework", "net20")
        $packagesConfig.DocumentElement.AppendChild($packageElement) | Out-Null
    }
    else
    {
        Write-Host "Updating current NuGet package version from" $packageElement.version "to $packageVersion"
        $packageElement.SetAttribute("version", $packageVersion)
    }


    return $postSharpNugetPath
    
}

function Backup-File
{
    param ( [string] $path )


    if ($backup)
    {
        $i = 0

        # Find a unique number.
        do { $i++}
        while ( Test-Path "$path.bak.$i" )

        Copy-Item $projectFullName "$path.bak.$i"
    }

}

function Upgrade-Project
{
    param(
        [string]$projectFullName,
        [string]$repositoryPath,
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
    $incompatibleProperty1 = $csproj.GetProperty('CustomAfterMicrosoftCompactFrameworkCommonTargets')
    $incompatibleProperty2 = $csproj.GetProperty('CreateSilverlightAppManifestDependsOn')
    if (($targetFrameworkIdentifier -and $targetFrameworkIdentifier.EvaluatedValue -ne '.NETFramework') -or $incompatibleProperty1 -or $incompatibleProperty2 )
    {
        Write-Warning "Project doesn't target .NET Framework. Skipping the project."
        return
    }

    # ignore if we have SkipPostSharp
    if ( $csproj.GetPropertyValue("SkipPostSharp") -eq "True" -or $csproj.GetPropertyValue("DefineConstants") -contains "SkipPostSharp" )
    {
         Write-Warning "SkipPostSharp detected. Skipping the project"
         return
    }

    # check if there are some references to PostSharp
    if ( !(Has-PostSharpReference( $csproj )) )
    {
        Write-Warning "No PostSharp reference. Skipping the project"
        return
    }

    $postSharpReferences = Get-DirectPostSharpReferences $csproj

    
    $toolkitReference = $postSharpReferences | Where-Object { $_.Include -like 'postsharp.toolkit*' }
    if ($toolkitReference -and $toolkitReference.lenght -ne 0)
    {
        Write-Warning "Project contains unsupported toolkit reference(s):"
        $toolkitReference | ForEach-Object { Write-Warning $_.Include }
        Write-Warning "Skipping the project."
        Write-Host '' 
        return
    }


    # Open packages.config
    $projectPath = Split-Path $projectFullName
    $packagesConfigPath = Join-Path $projectPath 'packages.config'


    $packagesConfig = Open-PackagesConfig -packageFullName $packagesConfigPath
    $packagePath = Install-Package -repositoryPath $repositoryPath -packageName 'PostSharp' -packageVersion $version -packagesConfig $packagesConfig

    $referenceGroup = $postSharpReferences[0].Parent

    $projectPath = Split-Path $projectFullName
    $relativePackagePath = Get-RelativePath $projectPath $packagePath
    $relativePostSharpTargetsPath = Join-Path $relativePackagePath 'tools\PostSharp.targets'
    
    # Remove elements from previous installations or versions.
    $nodesToRemove = @()
    $nodesToRemove += $csproj.Xml.Properties | Where-Object {$_.Name.ToLowerInvariant() -eq "dontimportpostsharp" }
    $nodesToRemove += $csproj.Xml.Imports | Where-Object {$_.Project.ToLowerInvariant().EndsWith("postsharp.targets") } 
    $nodesToRemove += $csproj.Xml.Targets | Where-Object {$_.Name.ToLowerInvariant() -eq "ensurepostsharpimported" }
    $nodesToRemove += $postSharpReferences

    $postSharpReferences | ForEach-Object { Write-Host "Removing reference" $_.Include }
    $nodesToRemove | ForEach-Object         {
          $nodeToRemove = $_
          $parent = $nodeToRemove.Parent
          $parent.RemoveChild($nodeToRemove) | out-null 

          # Remove the group if it becomes empty.
          if ( $parent.Count -eq 0 )
          {
            $parent.Parent.RemoveChild( $parent )
          }
          

        }

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


  
    # add reference to PostSharp.dll
    if ( $version -like "4.*" )
    {
        $libSubdir = "net35-client"
    }
    else
    {
        $libSubdir = "net20"
    }

    Add-Reference  $csproj $packagePath "lib\$libSubdir\PostSharp.dll"

    # install PostSharp.Settings.dll if we must
    $postsharpSettingsReference = $postSharpReferences | Where-Object { $_.Include -like 'postsharp.settings*' }
    if ($postsharpSettingsReference -and $postsharpSettingsReference.lenght -ne 0)
    {
        $postsharpSettingsPackagePath = Install-Package -repositoryPath $repositoryPath -packageName 'PostSharp.Settings' -packageVersion $version -packagesConfig $packagesConfig
        Add-Reference  $csproj $postsharpSettingsPackagePath 'lib\net40\PostSharp.Settings.dll'
    }


    # backup original file
    Backup-File $projectFullName

    # save the project file
    Checkout-File $outputProjectFullName
    $csproj.Save($outputProjectFullName)

    $packagesConfig.Save($packagesConfigPath)

}

function Open-PackagesConfig
{
    param(
        [string]$packageFullName
    )

   
    if (Test-Path $packageFullName)
    {
        $xml = [xml](Get-Content $packageFullName)

        Backup-File $packageFullName
    }
    else
    {
        Write-Host "Creating packages.config"
        $xml = [xml]"<packages></packages>"
    }


    return $xml    
}

function Upgrade-Directory
{
    param(
        [string]$rootPath = $null,
        [string]$postSharpVersion = $null,
        [string]$outputProjectFileNameSuffix = '',
        [bool]$backup = $true
    )
    $repositoryPath = Get-RepositoryPath $rootPath
    
    Write-Host ''
    Write-Host 'Updating projects'

    Get-ChildItem $rootPath -Recurse |
        Where-Object { ($_.Name -like '*.csproj') -or ($_.Name -like '*.vbproj') } |
        ForEach-Object {
            $project = $_
            try
            {
                Upgrade-Project -projectFullName $project.FullName -repositoryPath $repositoryPath -outputProjectFullName ($_.FullName + $outputProjectFileNameSuffix) -backup $backup -version $postSharpVersion
            }
            catch [Exception]
            {
                Write-Error ("Exception while processing " + $project.FullName)
                Write-Error $Error[0].Exception.ToString()
            }

            Write-Host ''
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
