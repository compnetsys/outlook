#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$filename = "salesforce_for_outlook.zip"
$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = $ENV:TMP + "\" + $filename
    $global:UNZIPTO=$env:tmp + "\salesforce_for_outlook"
    $global:reg = $wd + "\setform.reg"

    remove-Item $UNZIPTO -recurse -ErrorAction SilentlyContinue
    start-sleep 5
    New-Item $UNZIPTO -Type Directory



function get_salesforce()
    {
       IF(TEST-PATH  $DESTINATION){Remove-Item $DESTINATION}
         
        #Invoke-WebRequest $source -OutFile $destination

        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($SOURCE,$DESTINATION)
        
        #EXTRACT THE FILES
        
        $shell = new-object -com shell.application
        $zip = $shell.NameSpace(“$DESTINATION”)
        foreach($item in $zip.items())
        {
        $shell.Namespace(“$UNZIPTO”).copyhere($item, 0x14)
        }


    }
    
    

$outlookver = (get-itemproperty -literalpath HKLM:\SOFTWARE\Classes\Outlook.Application\CurVer).'(default)'
if($outlookver -like "*15*")
    {

        $makereg = (Get-Content $reg |  Foreach-Object {$_ -replace "14.0","15.0" } ) | set-content $reg -encoding ascii 
      

    }
    
  #IMPORT REGISTRY FILE
  cmd /c reg import $reg 2>&1 > $null  
  

$oftfile= $wd + "\Appointments.oft"
$vbfile = $wd + "\install_form.vbs"
$setformlocation = (get-content $vbfile | foreach-object{$_ -replace "formlocation","$oftfile"} ) | set-content $vbfile -encoding ascii

cmd /c c:\windows\system32\cscript.exe $vbfile