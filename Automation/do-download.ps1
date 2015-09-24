#DOWNLOAD LATEST VERSION TO THE TEMP FOLDER
$global:wd = split-path -parent $MyInvocation.MyCommand.Definition
$global:hostname = $env:computername

function get_salesforce()
    {


        $filename = "Salesforce_for_Outlook.zip"

        $SOURCE = "https://s3-us-west-2.amazonaws.com/compnetsys-software-delivery/" + $filename
        $DESTINATION = $ENV:TMP + "\" + $filename


        IF(TEST-PATH  $DESTINATION){Remove-Item $DESTINATION}
         
        #Invoke-WebRequest $source -OutFile $destination

        $DOWNLOAD = New-Object System.Net.WebClient
        $DOWNLOAD= $DOWNLOAD.DownloadFile($SOURCE,$DESTINATION)
        return


    }
    
    get_salesforce

