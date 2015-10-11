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

