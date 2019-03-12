#This script updates project files with an after build task and version files 
#in order to ensure assemblies are the same version as what is set in Jenkins
#Current as of 4/17/14:

param($installPath, $toolsPath, $package, $project) 

$dte.ExecuteCommand("File.SaveAll")

##############################
# Project Specific Variables #
#                            #
# Exit if Not Supported Type #
##############################

$proj = $project.FullName
$projType = $proj.SubString($proj.LastIndexOf(".") + 1)
$projFolder = Split-Path -parent $proj

if($projType -eq "csproj")
{
	$versionFile = "Properties\VersionInfo.cs"
	$versionFileContent = "using System.Reflection;`r`n`r`n[assembly: AssemblyVersion(`"1.0.0.0`")]"
	$assemblyFile = "Properties\AssemblyInfo.cs"
}
elseif($projType -eq "vbproj")
{
	$versionFile = "My Project\VersionInfo.vb"
	$versionFileContent = "Imports System.Reflection`r`n`r`n<assembly: AssemblyVersion(`"1.0.0.0`")>"
	$assemblyFile = "My Project\AssemblyInfo.vb"
}
elseif($projType -eq "vcxproj")
{
	$getFile = Get-ChildItem $projFolder\"*.rc" | select BaseName,Extension
	$versionFile = $GetFile.Basename + $getFile.Extension
	$assemblyFile = $versionFile
}
else
{
	Write-Host "Not a compatible project type, doing nothing"
	exit 0
}

############################
# Prepare Global variables #
############################

# Visual Studio Project Namespace
$ns = "http://schemas.microsoft.com/developer/msbuild/2003"

	#########################
	# Get Folder References #
	#########################
$currentDir = Get-Location
Set-Location $projFolder
$relativeInstallDir = Resolve-Path -Relative -Path $installPath
$buildVersioningLocation35 = Resolve-Path -Relative -Path ($toolsPath + "\..\lib\net35\Amcom.Build.Versioning.dll")
$buildVersioningLocation40 = Resolve-Path -Relative -Path ($toolsPath + "\..\lib\net40\Amcom.Build.Versioning.dll")
$assemblyFile = $projFolder + "\" + $assemblyFile
Set-Location $currentDir

########################################
# Add Build Versioning to Project File #
########################################


# Get the content of the proj file and cast it to XML for back end updating of the project file
$xml = [xml](get-content $proj)
$root = $xml.get_DocumentElement();
$mgr = new-object System.Xml.XmlNamespaceManager($xml.Psbase.NameTable)
$mgr.AddNamespace("ms", $ns)


		###########################
		# Check Before Build step #
		###########################
$BeforeBuild = $xml.SelectSingleNode("//ms:Project/ms:Target[@Name='BeforeBuild']", $mgr)
if(-not $BeforeBuild)
{
	$BeforeBuild = $xml.CreateElement("Target", $ns)
	$tname = $BeforeBuild.SetAttribute("Name", "BeforeBuild")
	$root.AppendChild($beforeBuild)
}
if($projType -eq "vcxproj") { $BeforeBuild.SetAttribute("BeforeTargets", "CLCompile") }
$BeforeBuild.SetAttribute("Condition", "Exists('`$(ProjectDir)$relativeInstallDir')")

$xml.Save($proj)

		##################################
		# Check for Jenkins Version Step #
		##################################

$JenkinsVersionStep = $xml.SelectSingleNode("//ms:Project/ms:Target[@Name='BeforeBuild']/ms:JenkinsVersion", $mgr)
if(-not $JenkinsVersionStep)
{
	$JenkinsVersionStep = $xml.CreateElement("JenkinsVersion", $ns)
	$BeforeBuild.AppendChild($JenkinsVersionStep)
}
$JenkinsVersionStep.SetAttribute("DestinationFile", $versionFile)

$xml.Save($proj)

		########################################
		# Create UsingJenkinsTask for .Net 4.0 #
		########################################

$UsingJenkinsTask40 = $xml.SelectSingleNode("//ms:Project/ms:UsingTask[@TaskName='JenkinsVersion'][1]", $mgr)
if(-not $UsingJenkinsTask40)
{
	$UsingJenkinsTask40 = $xml.CreateElement("UsingTask", $ns)
	$UsingJenkinsTask40.SetAttribute("TaskName", "JenkinsVersion")
	$root.AppendChild($UsingJenkinsTask40)
}
$UsingJenkinsTask40.SetAttribute("Condition", "'`$(TargetFrameworkVersion)' != 'v3.5' AND Exists('`$(ProjectDir)$buildVersioningLocation40')")
$UsingJenkinsTask40.SetAttribute("AssemblyFile", "`$(ProjectDir)$buildVersioningLocation40")

$xml.Save($proj)

		########################################
		# Create UsingJenkinsTask for .Net 3.5 #
		########################################

$UsingJenkinsTask35 = $xml.SelectSingleNode("//ms:Project/ms:UsingTask[@TaskName='JenkinsVersion'][2]", $mgr)
if(-not $UsingJenkinsTask35)
{
	$UsingJenkinsTask35 = $xml.CreateElement("UsingTask", $ns)
	$UsingJenkinsTask35.SetAttribute("TaskName", "JenkinsVersion")
	$root.AppendChild($UsingJenkinsTask35)
}
$UsingJenkinsTask35.SetAttribute("Condition", "'`$(TargetFrameworkVersion)' == 'v3.5' AND Exists('`$(ProjectDir)$buildVersioningLocation35')")
$UsingJenkinsTask35.SetAttribute("AssemblyFile", "`$(ProjectDir)$buildVersioningLocation35")

$xml.Save($proj)

		#########################
		# Create VersionInfo.cs #
		#########################

if($projType -eq "csproj" -or $projType -eq "vbproj")
{
	$CompileVersionFile = $xml.SelectSingleNode("//ms:Project/ms:ItemGroup/ms:Compile[@Include='$versionFile']", $mgr)
	if(-not $CompileVersionFile)
	{
		$itemGroup = $xml.CreateElement("ItemGroup", $ns)
		$CompileVersionFile = $xml.CreateElement("Compile", $ns)
		$CompileVersionFile.SetAttribute("Include", $versionFile)

		$itemGroup.AppendChild($CompileVersionFile)
		$root.AppendChild($itemGroup)
		
		$xml.Save($proj)
	}

	$versionFileFullPath = $projFolder + "\" + $versionFile
	if(-not (Test-Path ($versionFileFullPath)))
	{
		Write-Host $versionFileFullPath
		$temp = $versionFileFullPath + ".bak"
		Write-Host $temp
		New-Item $versionFileFullPath -type file
		Add-Content -value $versionFileContent -path $versionFileFullPath
		Get-Content $assemblyFile | Where-Object {$_ -notmatch "AssemblyVersion"} | Set-Content $temp
		Get-Content $temp | Where-Object {$_ -notmatch "AssemblyFileVersion"} | Set-Content $assemblyFile 
		Write-Host $assemblyFile
		Remove-Item $temp
	}
}



##############################################
# Remove Reference to Amcom.Build.Versioning #
##############################################

if($projType -eq "csproj" -or $projType -eq "vbproj")
{
	$node = $xml.SelectSingleNode("//ms:Project/ms:ItemGroup/ms:Reference[@Include='Amcom.Build.Versioning']", $mgr)
	if ($node) { [Void]$node.ParentNode.RemoveChild($node) }
	$xml.Save($proj)
}