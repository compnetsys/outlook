clear
#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername


$GLOBAL:filename = "salesforce_for_outlook_v2.exe"  #!IMPORTANT!!!!!!!! THIS HAS TO REFLECT THE NEW FILENAME UPLOADED TO S3
$GLOBAL:foldername = $filename.Substring(0,$filename.Length -4)
$GLOBAL:folderpath = "C:\windows\temp\" + $foldername +"\"
write-host $GLOBAL:folderpath


$global:SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
$global:DESTINATION = "C:\windows\temp\" + $filename
    $global:UNZIPTO = $DESTINATION
NEW-ITEM C:\TEMP -TYPE Directory -ErrorAction SilentlyContinue

function GLOBAL:GET_DATETIME()
{
    $GLOBAL:DATE = GET-DATE
    return $DATE

    }




$GLOBAL:LOGFILE = "C:\TEMP\SALESFORCE_INSTALL.LOG"
$GLOBAL:SFILOG = "c:\TEMP\SALESFORCELOG.TXT"
IF(TEST-PATH $LOGFILE ){REMOVE-ITEM $LOGFILE -ErrorAction SilentlyContinue}
$GLOBAL:LOGINIT = "INITIALIZING LOG FILE----------->$DATE`r" 
$GLOBAL:SANITIZE = "`tSANITIZING THE ENVIRONMENT`r"
$GLOBAL:SANITIZEEND = "`tSANITIZATION COMPLETED`r"
$GLOBAL:CHECKINSTALLED = "`tCHECKING FOR INSTALLED INSTANCE`r"
$GLOBAL:CHECKINSTALLEDEND = "`tCHECKING FOR INSTALLED INSTANCE END`r"
$GLOBAL:CHECKINSTALLEDAFTER = "`tCHECKING POST INSTALL STATUS`r"
$GLOBAL:CHECKINSTALLEDAFTEREND = "`tCHECKING POST INSTALL STATUS END`r"
$GLOBAL:MAKESTARTUP = "`tMAKING NECESSARY STARTUP ENTRIES`r"
$GLOBAL:MAKESTARTUPEND = "`tCOMPLETED MAKING STARTUP ENTRIES`r"
$GLOBAL:DODOWNLOAD ="`tBEGINNING DOWNLOAD-----$DATE-`r"
$GLOBAL:DODOWNLOADNO = "`t`tDOWNLOAD NOT NEEDED - ALREADY AT LATEST FILE`r"
$GLOBAL:DODOWNLOADEND ="`t`tCOMPLETED DOWNLOAD-----$DATE-`r"
$GLOBAL:DODOWNLOADEXIT ="`tDOWNLOAD OF $SOURCE  FAILED  EXITING-----$DATE-`r`n`tMAKE SURE THAT THE FILE IS SET TO PUBLIC ON AMAZON S3 STORAGE OR YOU HAVE AMPLE STORAGE ON THE C DRIVE"
$GLOBAL:DOINSTALL = "`tBEGINNING SOFTWARE INSTALL ----- $DATE`r"
$GLOBAL:DOINSTALLEND = "`t`tSOFTWARE INSTALL COMPLETED --- $DATE`r"
$GLOBAL:DOSENDEMAIL = "`t`tPREPARING TO SEND EMAIL ----$DATE`r"
$GLOBAL:DOSENDEMAILEND = "`t`tCOMPLETED SENDING EMAIL ----$DATE`r"
$GLOBAL:DOSKIPDOTNET = "`t`tSKIPPING DOT NET INSTALL"


ECHO $LOGINIT | OUT-FILE $LOGFILE -Encoding unicode -Append


function sanitize()
{
  ECHO $SANITIZE  | OUT-FILE $LOGFILE -Encoding unicode -Append
  cmd /c rd /s /q c:\windows\temp 2>&1 > $null
  ECHO $SANITIZEEND | OUT-FILE $LOGFILE -Encoding unicode -Append
}

function check_installed()
{

    ECHO $CHECKINSTALLED  | OUT-FILE $LOGFILE -Encoding unicode -Append
  $file = "C:\Program Files (x86)\salesforce.com\Salesforce for Outlook\SfdcMsOl.exe"
  $formpath = "C:\ProgramData\graebel_outlook_form"
    if((Test-Path $file) -and (Test-path $formpath))
    {
    $GLOBAL:msg ="SALESFORCE ALREADY INSTALLED ---NOTHING TO DO HERE ON $HOSTNAME"
    #write-host "SALESFORCE ALREADY INSTALLED ---NOTHING TO DO HERE ON $HOSTNAME"
    $GLOBAL:SUBJECT = "Salesforce alread installed on $hostname"
    send_email
    break
    }

    ECHO $CHECKINSTALLEDEND| OUT-FILE $LOGFILE -Encoding unicode -Append
}


function check_installed_after()
{
  
  ECHO $CHECKINSTALLEDAFTER | OUT-FILE $LOGFILE -Encoding unicode -Append
  $file = "C:\Program Files (x86)\salesforce.com\Salesforce for Outlook\SfdcMsOl.exe"
  $formpath = "C:\ProgramData\graebel_outlook_form"
    if(Test-Path -path $file)
    {
    $GLOBAL:msg ="INSTALLED SUCCESSFULLY ---ON $HOSTNAME"
    $GLOBAL:SUBJECT = "Salesforce installed successfully on $HOSTNAME "
    }
    ELSE
    {

    $GLOBAL:msg ="SALESFORCE INSTALL FAILED ON $HOSTNAME. PLEASE REVIEW THE TASK INFORMATION IN GFI $HOSTNAME TO DETERMINE WHAT WENT WRONG. "
    $GLOBAL:SUBJECT = "SALESFORCE INSTALL FAILED ON $HOSTNAME."
    }
    ECHO $CHECKINSTALLEDAFTEREND| OUT-FILE $LOGFILE -Encoding unicode -Append
    send_email
}


function make_startup()
{
   
   ECHO $MAKESTARTUP | OUT-FILE $LOGFILE -Encoding unicode -Append
   #CREATE THE NECESSARY SHORTCUTS TO ALLOW STARTUP ON USER LOGIN


  $ShortcutFile = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Salesforce Form Install.lnk" 
  if(Test-Path $shortcutfile){Remove-Item $shortcutfile}

  $COPYPATH = "$folderpath" + "graebel_outlook_form\"
  
  Copy-Item $COPYPATH "C:\programdata\" -recurse -force

    
  $shortcuttarget = "C:\ProgramData\graebel_outlook_form\Form_Install_Launcher.cmd"
  

    $TargetFile = "$shortcuttarget"
    $WScriptShell = New-Object -ComObject WScript.Shell
    $Shortcut = $WScriptShell.CreateShortcut($ShortcutFile)
    $Shortcut.TargetPath = $TargetFile
    $Shortcut.Save()

$rule=new-object System.Security.AccessControl.FileSystemAccessRule ("everyone","FullControl","Allow")            
$file = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\Startup\Salesforce Form Install.lnk"

          
    $acl = Get-ACL -Path $File -ErrorAction stop            
    $acl.SetAccessRule($rule)            
    Set-ACL -Path $File -ACLObject $acl -ErrorAction stop            
         
    ECHO $MAKESTARTUPEND | OUT-FILE $LOGFILE -Encoding unicode -Append

}


function do_download($src, $dst)
    {
       
       ECHO $DODOWNLOAD | OUT-FILE $LOGFILE -Encoding unicode -Append
        
       IF(TEST-PATH  $dst)
       {
            ECHO $DODOWNLOADNO | OUT-FILE $LOGFILE -Encoding unicode -Append

            }

            else
            {
        
        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($src,$dst)

        }
        
        #EXTRACT THE FILES
       $global:CMDUNZIP = "C:\WINDOWS\TEMP\UNZIP.CMD"
       echo "$UNZIPTO" | out-file $CMDUNZIP -Encoding ASCII
       start-process cmd -Argumentlist "/c $CMDUNZIP" -workingdirectory C:\windows\temp -nonewwindow -wait
       
        ECHO $DODOWNLOADEND | OUT-FILE $LOGFILE -Encoding unicode -Append
    }
       
function check_download()
{
    if(!(Test-Path $DESTINATION))
    {
        ECHO $DODOWNLOADEXIT | OUT-FILE $LOGFILE -Encoding unicode -Append
        send_email
        break
        }
        

}

function install_salesforce()
{
   
   
   ECHO $DOINSTALL | OUT-FILE $LOGFILE -Encoding unicode -Append

    #write-host $user
        #------------CRYPTOGRAPHY STUFF-----######
  $enccode = "76492d1116743f0423413b16050a5345MgB8AFkASwBGADYANAB4AHkAdwB5AEkAYgBsAFIAVwA2AEYAegAvAFIATABvAHcAPQA9AHwAMQBiADcANQA5ADEAMABiAGYAYwA5ADQAYwA4ADcANgBhADgANwBlAGMANAAwAGMAMQA3ADcAOQBhADIAZQBiAA=="

        $key = (29,198,19,130,110,209,124,187,56,144,7,99,79,240,71,85)
        $username = "$hostname\administrator"
        $EncryptedPW = $enccode
        $SecureString = ConvertTo-SecureString -String $EncryptedPW -Key $Key
        $loccred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username,$SecureString
        #------------END CRYPTO--------------#####
     
               
          #Start-Process powershell -Credential $loccred -ArgumentList '-noprofile -command &{Start-Process C:\windows\temp\salesforce_for_outlook\do_install.cmd -verb runas }'
          #cmd /c C:\windows\temp\salesforce_for_outlook\do_install.cmd
                
                #UNINSTALL VTSOT VERSION 10.06
                $killoutlook = get-process  | where-object {$_.processname -eq "outlook"}
                    if($killoutlook)
                    {
                        stop-process -name outlook -force
    
                        }

                $killmsi = get-process  | where-object {$_.processname -eq "msiexec"}
                    if($killmsi)
                    {
                        stop-process -name msiexec -force
                        
    
                        }

                      
                start-process -FilePath "c:\windows\system32\msiexec.exe"  -ArgumentList "/X{7C0242A3-8B66-35D1-9FE0-13B426ACB609} /quiet /norestart" -wait

             


                $vstorexe = $folderpath + "vstor_redist.exe"
                start-process -filepath "$vstorexe" -ArgumentList "/q /norestart /log C:\temp\VSTOR.log" -Wait

                #IF DOT NET 4 IS INSTALLED THEN SKIP THE INSTALLATION
                $getdotnetv = (Get-ChildItem -Path $Env:windir\Microsoft.NET\Framework | 
                Where-Object {$_.PSIsContainer -eq $true } |
                Where-Object {$_.Name -match 'v\d\.\d'} | 
                Sort-Object -Property Name -Descending | 
                Select-Object -First 1).Name

                 if($getdotnetv -like "*v4.5*")
                 {
                     ECHO $DOSKIPDOTNET| OUT-FILE $LOGFILE -Encoding unicode -Append
                  }
                  else
                  {
                        $dotnetexe = $folderpath + "NDP452-KB2901907-x86-x64-AllOS-ENU.exe"
                        Start-Process -FilePath "$dotnetexe" -ArgumentList "/q /norestart" -Wait

                 }

              

                $o2010exe = $folderpath + "o2010pia.msi"
                start-process -filepath "$o2010exe" -argumentlist "/quiet /norestart" -Wait 

                $vcredexe = $folderpath + "vcredist_x86.exe"
                start-process -filepath "$vcredexe" -argumentlist  "/q /norestart" -Wait

                start-sleep -seconds 240


                $killoutlook = get-process  | where-object {$_.processname -eq "outlook"}
                    if($killoutlook)
                    {
                        stop-process -name outlook -force
    
                        }

                $sfexe = $folderpath + "setup.x64.msi"
                start-process -FilePath "c:\windows\system32\msiexec.exe" -ArgumentList " /i $sfexe /quiet /norestart /l*v C:\temp\salesforcelog.txt" -Wait


   ECHO $DOINSTALLEND | OUT-FILE $LOGFILE -Encoding unicode -Append
}


function global:send_email()
{


ECHO $DOSENDEMAIL | OUT-FILE $LOGFILE -Encoding unicode -Append
$CredUser = "alerts@computernetworksystems.us"
$CredPassword = "stlucianice"

$EmailFrom = "salesforceinstall@graebelmoving.com"
$EmailTo = "salesforceinstalls@compnetsys.com" 
$Subject = "$subject"
$Body = "$msg"

$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
$attachment = New-Object System.Net.Mail.Attachment($LOGFILE)
$SMTPMessage.Attachments.Add($attachment)
if(Test-Path $SFILOG)
{
$attachment2 = New-Object System.Net.Mail.Attachment($SFILOG)
$SMTPMessage.Attachments.Add($attachment2)
}



$SMTPServer = "smtpout.secureserver.net" 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 80) 
$SMTPClient.EnableSsl = $false
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($CredUser, $CredPassword); 
#$SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body)
$SMTPClient.Send($SMTPMessage)
#ECHO $DOSENDEMAILEND | OUT-FILE $LOGFILE -Encoding unicode -Append

}


function update_form()
{

<#
SYNOPSIS :
    This script will install a custom user form in microsoft outlook. This form was requested by Salesforce for integration with microsoft outlook"
    
    Written by the Team at comptuer Network systems : wwww.compnetsys.com
    
  Mckin S
  Melvin F
  
  
 ----#>  
    


#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername
$global:BASESOURCE = "https://s3.amazonaws.com/gvl/sf-form/"
    $global:formpath = "C:\ProgramData\graebel_outlook_form"
    $purgefolder = $formpath + "\*"
    Remove-Item $purgefolder -ErrorAction SilentlyContinue
    New-Item $formpath -Type Directory -ErrorAction SilentlyContinue



      $global:file1 = "form-install.ps1"
      $global:file2 = "Salesforce-Form-Install.lnk"
      $global:file3 =  "form_install_launcher.cmd"



            function global:download($src,$dst)
                {
    
                   WRITE-HOST "`r`nDOWNLOADING LATEST FORM FILES $src ----->`r`n"
                   IF(TEST-PATH  $dst){Remove-Item $dst}
         
                   #Invoke-WebRequest $source -OutFile $destination

                   $DOWNLOAD = New-Object System.Net.WebClient
                   $DOWNLOAD= $DOWNLOAD.DownloadFile($src,$dst)
           

                   if(Test-Path $dst) { WRITE-HOST "`tDOWNLOAD COMPLETED FOR $src"}
                }





function get_files()

{

      

            $global:filelist = ("$file1","$file2","$file3")
            $global:getfiles = {$filelist}.Invoke()


            foreach($file in $getfiles)
                {
                    $source = $basesource + "$file"
                    $destination = $formpath + "\$file"
                        remove-Item $destination -recurse -ErrorAction SilentlyContinue


                    download $source $destination


                }
}

function make_shortcut()
{

   $global:ShortcutFile = $formpath + "\Salesforce-Form-Install.lnk" 
   $desktop = ([Environment]::GetEnvironmentVariable("Public"))+"\Desktop"
   Copy-Item $shortcutfile $desktop -Force
  

$rule=new-object System.Security.AccessControl.FileSystemAccessRule ("everyone","FullControl","Allow")            
$file = $desktop + "\Salesforce-Form-Install.lnk"

          
    $acl = Get-ACL -Path $File -ErrorAction stop            
    $acl.SetAccessRule($rule)            
    Set-ACL -Path $File -ACLObject $acl -ErrorAction stop            
         
 

}


function global:send_email()
{


$CredUser = "alerts@computernetworksystems.us"
$CredPassword = "stlucianice"

$EmailFrom = "salesforceinstall@graebelmoving.com"
$EmailTo = "salesforceinstalls@compnetsys.com" 
$Subject = "$subject"
$Body = "$msg"

$SMTPMessage = New-Object System.Net.Mail.MailMessage($EmailFrom,$EmailTo,$Subject,$Body)
#$attachment = New-Object System.Net.Mail.Attachment($LOGFILE)
#$SMTPMessage.Attachments.Add($attachment)
#if(Test-Path $SFILOG)
#{
#$attachment2 = New-Object System.Net.Mail.Attachment($SFILOG)
#$SMTPMessage.Attachments.Add($attachment2)
#}



$SMTPServer = "smtpout.secureserver.net" 
$SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer, 80) 
$SMTPClient.EnableSsl = $false
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential($CredUser, $CredPassword); 

$SMTPClient.Send($SMTPMessage)


}

clear
get_files
make_shortcut

$testfile1 = $shortcutfile
$testfile2 = $formpath + "\"+$file1 

if((Test-Path $testfile1) -and (Test-Path $testfile2))
{
    $global:subject = "SUCCEEDED : SALESFORCE FORM $HOSTNAME"
    send_email

}

else
{


$global:subject = "FAILED : SALESFORCE FORM FILES $HOSTNAME"
    send_email


}

}


#sanitize
do_download $SOURCE $DESTINATION
update_form
check_download
#make_startup
check_installed
install_salesforce

check_installed_after

