#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$filename = "salesforce_for_outlook.zip"
$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = $ENV:TMP + "\" + $filename
    $global:UNZIPTO=$env:tmp + "\salesforce_for_outlook"

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
    
function find_os()
    {
        
        if(!(Test-Path "C:\Program Files (x86)"))
        {
            $global:exe  = $env:tmp + "\salesforce_for_outlook\install\setup.msi"
            
          }
          else
         {
             $global:exe = $env:tmp + "\salesforce_for_outlook\install\setup.x64.msi"
          }
          
    }

get_salesforce
find_os

write-host $exe

function install_salesforce()
{
    cmd /c  taskkill /im outlook.exe 2>&1 > $null
    start-sleep 10
    
    $command = --% "$exe /quiet /norestart ALLUSERS="1" "
    cmd /c $command

}


install_salesforce
