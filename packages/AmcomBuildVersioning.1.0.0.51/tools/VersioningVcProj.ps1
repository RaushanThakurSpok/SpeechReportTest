#This script updates vc project files in VS 2008 
#It is called from command line as a pre-build event in order to take
#in Jenkins version variables and apply them to the project so that 
#all of the assemlby versions match the product version
#Current as of 7/26/13:

#parameter passed in from command line
param([string]$projectPath)

#Quotes are needed to ensure the full project path is used.
#However, if a path does not need quotes, the quotes will also be passed in.
#This line removes such a case.
$projectPath = $projectPath -replace '"', ""


#Get the Version Resource file name so that we can modicy it later
$GetFile = Get-ChildItem ("{0}\*.rc" -f $projectPath) | select BaseName,Extension
$VersionFile = $GetFile.Basename + $GetFile.Extension

#Check if the Version Environment Variables exist, if they do, use that value, otherwise, use default values
If (Test-Path env:MAJOR_VERSION)
{
	$MAJOR = (Get-Item env:MAJOR_VERSION).value
}
else
{
	$MAJOR = "1"
}
If (Test-Path env:MINOR_VERSION)
{
	$MINOR = (Get-Item env:MINOR_VERSION).value
}
else
{
	$MINOR = "0"
	}
If (Test-Path env:REVISION)
{
	$REVISION = (Get-Item env:REVISION).value
	}
else
{
	$REVISION = "0"
}
If (Test-Path env:BUILD_NUMBER)
{
	$BUILD = (Get-Item env:BUILD_NUMBER).value
}
else
{
	$BUILD = "0"
}


#Updated how sections are replaced to provide an easy means of updating this section.
$ReplaceList = @{
"(?-i)FILEVERSION(.*)" = "FILEVERSION $MAJOR,$MINOR,$REVISION,$BUILD"
'(?-i)FileVersion"(.*)' = "FileVersion"", ""$MAJOR, $MINOR, $REVISION, $BUILD"""
"(?-i)PRODUCTVERSION(.*)" = "PRODUCTVERSION $MAJOR,$MINOR,$REVISION,$BUILD"
'(?-i)ProductVersion"(.*)' = "ProductVersion"", ""$MAJOR, $MINOR, $REVISION, $BUILD"""
}

#Create a variable to hold the ResourceFile reference.
#This is important as it provides a single location to change how paths are used.
$resourceFile = "{0}\{1}" -f $projectPath, $VersionFile

#Updated Versioning Script to Get all the content and close the file in one step before working on replacing data.
#This section will Read data from Resource file
#Iterate line-by-line and replace any matches to the keys of "ReplaceList"
#Saves the data back to the resource file.
(Get-Content ($resourceFile)) |
  Foreach-Object {
    $line = $_
    $ReplaceList.GetEnumerator() | ForEach-Object {
        if ($line -match $_.Key) { $line = $line -replace $_.Key, $_.Value }
    }
    $line
  } | Set-Content ($resourceFile)