Function Write-TextToSpeech
{
<#
.SYNOPSIS
Converts text to speech

.DESCRIPTION
The Write-TextToSpeech function speaks the specified content with either an installed female or male voice, which is based on the Microsoft .NET Framework SpeechSynthesizer class.
The minumum version of PowerShell required is version 5.0, since it uses the constructor [System.Speech.Synthesis.SpeechSynthesizer]::New() to instantiate the $objSpeech object.

.EXAMPLE
Write-TextToSpeech -Text "Hello World" -Gender Female
Says "Hello World" using the default sound device and the installed female voice if available.

.EXAMPLE
Write-TextToSpeech -Text "Hello World" -Gender Male
Says "Hello World" using the default sound device and the installed male voice if available.

.PARAMETER Text
The text to be spoken.

.PARAMETER Gender
The gender of the voice used to speak the text provided. 

.INPUTS
[string]

.OUTPUTS
[SpeechSynthesizer]

.NOTES
NAME: Write-TextToSpeech

REQUIREMENTS: 
-Version 5.0

AUTHOR: Preston K. Parsard

ATTRIBUTION: 
The inspiration for this script originated from Guido Basilio de Oliveira, an MVP and Microsoft Parnter. More information about Guido is available at:
https://social.technet.microsoft.com/profile/guido%20basilio%20de%20oliveira/. The link to his original script is also provided below in the .LINK section of this header.

LASTEDIT: 12 NOV 2016

KEYWORDS: Speech, Voice

LICENSE:
The MIT License (MIT)
Copyright (c) 2016 Preston K. Parsard

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), 
to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software. 

DISCLAIMER: 
THIS SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, 
royalty-free right to use and modify the Sample Code and to reproduce and distribute the Sample Code, provided that You agree: (i) to not use Our name, 
logo, or trademarks to market Your software product in which the Sample Code is embedded; 
(ii) to include a valid copyright notice on Your software product in which the Sample Code is embedded; and (iii) to indemnify, hold harmless, 
and defend Us and Our suppliers from and against any claims or lawsuits, including attorneys’ fees, 
that arise or result from the use or distribution of the Sample Code.

.LINK
https://www.powershellgallery.com/
https://gallery.technet.microsoft.com/scriptcenter/Out-Speech-298142d3
https://msdn.microsoft.com/en-us/library/windows/apps/windows.media.speechsynthesis.speechsynthesizer.aspx
https://msdn.microsoft.com/en-us/library/system.speech.synthesis.speechsynthesizer(v=vs.110).aspx 
#>

<# WORK ITEMS
TASK-INDEX: 000
#>

<#
**************************************************************************************************************************************************************************
REVISION/CHANGE RECORD	
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
DATE         VERSION    NAME			     CHANGE
--------------------------------------------------------------------------------------------------------------------------------------------------------------------------
12 NOV 2016  01.0.00.00 Preston K. Parsard Initial release
#>

 [CmdletBinding()]
 Param
 (
  [Parameter(Mandatory=$true, ValueFromPipeline=$true)][String] $Text,
  [ValidateSet('Male','Female')][String] $Gender
 ) #end Param

 # Initialize automatic error variable
 $Error.Clear()

 # Attempt to load the required assembly
 Try { Add-Type -AssemblyName System.Speech -ErrorAction Stop }

 # If an error was encountered loading the assembly, inform user and provide addtional references.
 Catch 
 { 
  Clear-Host
  $CatchMessage = 
@"
The required components are not available on this system. 
For more information about the requirements and details of the .NET Framework SpeechSynthesizer Class, please see the following:
1. https://msdn.microsoft.com/en-us/library/windows/apps/windows.media.speechsynthesis.speechsynthesizer.aspx
2. https://msdn.microsoft.com/en-us/library/system.speech.synthesis.speechsynthesizer(v=vs.110).aspx 
Terminating Write-TextToSpeech function...
"@ 
 Write-Host $CatchMessage -ForegroundColor Yellow
 } # end Catch

# Reset the automatic error variable
Finally 
{ $Error.Clear() } 

$objSpeech = [System.Speech.Synthesis.SpeechSynthesizer]::New()
$objSpeech.SelectVoiceByHints($Gender)
$objSpeech.Speak($Text)
} #end Function