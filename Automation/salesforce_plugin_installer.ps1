#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$filename = "salesforce_for_outlook.zip"
$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = $ENV:TMP + "\" + $filename
    $global:UNZIPTO=$env:tmp + "\salesforce_for_outlook"

    remove-Item $UNZIPTO -recurse -ErrorAction SilentlyContinue
    start-sleep 5
    New-Item $UNZIPTO -Type Directory -Erroraction Silentlycontinue



function do_download($src, $dst , $zipto , $dozip)
    {
       IF(TEST-PATH  $dst){Remove-Item $dst}

        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($src,$dst)
        
        #EXTRACT THE FILES
        if($dozip -eq "yes")
        {
            $shell = new-object -com shell.application
            $zip = $shell.NameSpace(“$dst”)
            foreach($item in $zip.items())
            {
            $shell.Namespace(“$zipto”).copyhere($item, 0x14)
            }
        }


    }
       
function exe_paths()
    {
        
        if(!(Test-Path "C:\Program Files (x86)"))
        {
            $global:exe  = $env:tmp + "\salesforce_for_outlook\install\setup.msi"
            $global:pia  = $env:tmp + "\salesforce_for_outlook\install\o2010pia.msi"
            $global:dotnet = $env:tmp + "\salesforce_for_outlook\install\dotNetFx40_Full_x86_x64.exe"
             $global:vc = $env:tmp + "\salesforce_for_outlook\install\vcredist_x86.exe"
             $global:vstore = $env:tmp + "\salesforce_for_outlook\install\vstor_redist.exe"
            
          }
          else
         {
             $global:exe = $env:tmp + "\salesforce_for_outlook\install\setup.x64.msi"
             $global:pia  = $env:tmp + "\salesforce_for_outlook\install\o2010pia.msi"
             $global:dotnet = $env:tmp + "\salesforce_for_outlook\install\dotNetFx40_Full_x86_x64.exe"
              $global:vc = $env:tmp + "\salesforce_for_outlook\install\vcredist_x86.exe"
              $global:vstore = $env:tmp + "\salesforce_for_outlook\install\vstor_redist.exe"
          }
          
    }




function install_salesforce()
{
   
   
    #

    #write-host $user
        #------------CRYPTOGRAPHY STUFF-----######
  $enccode = "76492d1116743f0423413b16050a5345MgB8AFkASwBGADYANAB4AHkAdwB5AEkAYgBsAFIAVwA2AEYAegAvAFIATABvAHcAPQA9AHwAMQBiADcANQA5ADEAMABiAGYAYwA5ADQAYwA4ADcANgBhADgANwBlAGMANAAwAGMAMQA3ADcAOQBhADIAZQBiAA=="

        $key = (29,198,19,130,110,209,124,187,56,144,7,99,79,240,71,85)
        $encfile =  "$wd" + "\creds\enccreds_user.txt"
        $username = "$hostname\administrator"
        $EncryptedPW = $enccode
        $SecureString = ConvertTo-SecureString -String $EncryptedPW -Key $Key
        $loccred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$SecureString
        #------------END CRYPTO--------------#####
     
 
            $msiexec = "C:\windows\system32\msiexec.exe"
            
             $commandvs = "/q /norestart"
            start-process $vstore $commandvs  -Credential $loccred -loaduserprofile 
            
            $commandvc = "/install /quiet /norestart"
            start-process $vc $commandvc  -Credential $loccred -loaduserprofile 
            
            $commanddotnet = "/q /norestart"
            start-process $dotnet $commanddotnet  -Credential $loccred -loaduserprofile 
            
            $commandpia = "/package $exe /quiet /norestart "
           start-process "$msiexec"  $commandpia  -Credential $loccred -loaduserprofile 
            
            cmd /c  taskkill /im outlook.exe 2>&1 > $null
            start-sleep 10
         
           $commandexe = "/package $exe /quiet /norestart "
            start-process "$msiexec"  $commandexe  -Credential $loccred -loaduserprofile 
   
       

}

do_download $SOURCE $DESTINATION $UNZIPTO "YES"
exe_paths
install_salesforce
