<# ----
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
    $global:TMP=$env:tmp



      $global:file1 = "Appointments.oft"
      $global:file2 = "Install_form.vbs"
      $global:file3 = "setform.reg"

       $global:oftfile= $TMP + "\" + $file1
       $global:vbfile = $TMP + "\" + $file2
       $global:reg = $TMP + "\" + $file3



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
                    $destination = $TMP + "\$file"
                        remove-Item $destination -recurse -ErrorAction SilentlyContinue


                    download $source $destination


                }
}


    function install_form()
    
    {
        WRITE-HOST "`r`nINSTALLING FORM ----->"
        
        #DETERMING VERSION OF OUTLOOK INSTALLED
        $outlookver = (get-itemproperty -literalpath HKLM:\SOFTWARE\Classes\Outlook.Application\CurVer).'(default)'
        if($outlookver -like "*15*")
         {

              $makereg = (Get-Content $reg |  Foreach-Object {$_ -replace "14.0","15.0" } ) | set-content $reg -encoding ascii 
              

          }
            
        #IMPORT REGISTRY FILE
        cmd /c reg import $reg 2>&1 > $null  
          

      
        $setformlocation = (get-content $vbfile | foreach-object{$_ -replace "formlocation","$oftfile"} ) | set-content $vbfile -encoding ascii

        #IMPORT FORM INTO OUTLOOK
        cmd /c c:\windows\system32\cscript.exe $vbfile
        
        WRITE-HOST "`tFORM INSTALL COMPLETED. THIS WINDOW WILL CLOSE IN 20 SECONDS`r`n"  
     }
        
        
        
 clear
   
 get_files
 install_form
 
 start-sleep 20
 EXIT
 
 
 