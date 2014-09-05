param(
    $solutionRootPath = $null,
    $nugetVersion = $null,
    $outputProjectFileNameSuffix = '',
    $backup = $true
)

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

function Get-RepositoryPath($solutionRootPath)
{
    $repositoryPath = .\NuGet.exe config repositorypath
    if ($repositoryPath -like 'WARNING*')
    {
        Write-Warning 'repositorypath nuget setting is not set, using solution path as root for repository path'
        $repositoryPath = Join-Path $solutionRootPath 'packages'
    }

    Write-Host "Using $repositorypath as repositorypath"

    return $repositoryPath
}

function Upgrade-Project
{
    param(
        $projectFullName,
        $packagePath,
        $outputProjectFullName = $null,
        $backup = $true
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

    # check if there are some references to PostSharp
    $postSharpReferences = Get-PostSharpReference $csproj

    if (!$postSharpReferences -or $postSharpReferences.length -eq 0)
    {
        Write-Warning "No PostSharp reference in $projectFullName. Skipping the project"
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

    if ($nodesToRemove -and $nodesToRemove.length)
    {
        $nodesToRemove | ForEach-Object { $_.Parent.RemoveChild($_) | out-null }
    }

    # Set property DontImportPostSharp to prevent locally-installed previous versions of PostSharp to interfere.
    $csproj.Xml.AddProperty( "DontImportPostSharp", "True" ) | Out-Null

    # Add import to PostSharp.targets
    $import = $csproj.Xml.AddImport($relativePostSharpTargetsPath)
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
        Copy-Item $projectFullName ($projectFullName + ".backup")
    }

    # save the project file
    $csproj.Save($outputProjectFullName)
}

function Upgrade-Solution
{
    param(
        $solutionRootPath,
        $nugetVersion,
        $outputProjectFileNameSuffix = '',
        $backup = $true
    )

    try
    {
        $repositoryPath = Get-RepositoryPath $solutionRootPath
        $nugetPackage = 'PostSharp.' + $nugetVersion
    
        .\NuGet.exe install 'PostSharp' -Version $nugetVersion -OutputDirectory $repositoryPath
    
        $postSharpNugetPath = Join-Path $repositoryPath $nugetPackage

        Write-Host ''
        Write-Host 'Updating projects'

        Get-ChildItem $solutionRootPath -Recurse |
            Where-Object { $_.Name -like '*.csproj' } |
            ForEach-Object { Upgrade-Project -projectFullName $_.FullName -packagePath $postSharpNugetPath -outputProjectFullName ($_.FullName + $outputProjectFileNameSuffix) -backup $backup }
    }
    catch [Exception]
    {
        Write-Error $_.Message
        Write-Error 'Unhandled exception thrown. Terminating batch upgrade.'
    }
}

#Upgrade-Project 'C:\src\upgrade\test\projects\Core\PostSharp.CommandLine.Cil\PostSharp.CommandLine.Cil-x64-4.0.csproj' 'C:\src\upgrade\test\projects\Core\PostSharp.CommandLine.Cil\PostSharp.CommandLine.Cil-x64-4.0.csproj.new' 'c:\src\upgrade\test\projects\packages\PostSharp.3.2.27-beta\'

#Upgrade-Solution -solutionRootPath 'C:\src\upgrade\test\projects' -nugetVersion '3.2.27-beta' -outputProjectFileNameSuffix '.new' -backup $false

if ($solutionRootPath -and $nugetVersion)
{
    Upgrade-Solution -solutionRootPath $solutionRootPath -nugetVersion $nugetVersion -outputProjectFileNameSuffix $outputProjectFileNameSuffix -backup $backup
}
