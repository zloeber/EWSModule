function Get-EWSMailboxAttachments {
    [CmdletBinding()] 
    param( 
        [Parameter(Position=0, Mandatory=$true)] [string]$MailboxName,
        [Parameter(Position=1, Mandatory=$true)] [PSCredential]$Credentials,
        [Parameter(Position=2, Mandatory=$false)] [PSObject]$AttachmentList,
        [Parameter(Position=3, Mandatory=$false)] [string]$AttachmentType,
        [Parameter(Position=4, Mandatory=$false)] [string]$AttachmentName,
        [Parameter(Position=5, Mandatory=$false)] [switch]$Download,
        [Parameter(Position=6, Mandatory=$false)] [string]$DownloadDirectory
    )  
    $KQL = ""
    if([string]::IsNullOrEmpty($DownloadDirectory)){
        $DownloadDirectory = (Get-Location).Path
    }
    $service = Connect-Exchange -MailboxName $MailboxName -Credential $Credentials
    if([string]::IsNullOrEmpty($AttachmentType) -band [string]::IsNullOrEmpty($AttachmentName)){
        if([string]::IsNullOrEmpty($AttachmentList))
        {
            $AttachmentList = @()
            $AttachmentList += "xlsx"
            $AttachmentList += "docx"
            $AttachmentList += "doc"
            $AttachmentList += "xls"
            $AttachmentList += "pptx"
            $AttachmentList += "ppt"
            $AttachmentList += "txt"
            $AttachmentList += "mp3"
            $AttachmentList += "zip"
            $AttachmentList += "txt"
            $AttachmentList += "wma"
            $AttachmentList += "pdf"
        }
    }
    else{
        $AttachmentList = @()
        if(![string]::IsNullOrEmpty($AttachmentType)){
            $AttachmentList += $AttachmentType
        }
        else{
            $KQL = "Attachment:" + $AttachmentName
        }
    }    
    if($KQL -eq ""){
        foreach($Item in $AttachmentList){
            if($KQL -eq ""){
                $KQL = "Attachment:." + $Item
            }
            else{
                $KQL += " OR Attachment:." + $Item
            }
        }            
    }
    if($Download)
    {
        $attachemtItems = New-EWSeDiscoveryPreviewItems -service $service -KQL $KQL -SearchableMailboxString $MailboxName 
        foreach($AttachmentItem in $attachemtItems){
            Write-Host ("Processing Item " + $AttachmentItem.Subject)
            $Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$AttachmentItem.Id)
            foreach($attach in $Item.Attachments){
                if($attach -is [Microsoft.Exchange.WebServices.Data.FileAttachment]){
                    if(![string]::IsNullOrEmpty($attach.Name) -band $attach.IsInline -eq $false){
                        $attach.Load()    
                        $FileName = New-UniqueFileName -FileName ($DownloadDirectory + “\” + $attach.Name.ToString())
                        $fiFile = new-object System.IO.FileStream($FileName, [System.IO.FileMode]::Create)
                        $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                        $fiFile.Close()
                        write-host ("Downloaded Attachment : " + $FileName)
                    }
                }
                if ("Microsoft.Exchange.WebServices.Data.ReferenceAttachment" -as [type]) {
                    if($attach -is [Microsoft.Exchange.WebServices.Data.ReferenceAttachment]){
                        $SharePointClientDll = (($(Get-ItemProperty -ErrorAction SilentlyContinue -Path Registry::$(Get-ChildItem -ErrorAction SilentlyContinue -Path 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SharePoint Client Components\'|Sort-Object Name -Descending| Select-Object -First 1 -ExpandProperty Name)).'Location') + "ISAPI\Microsoft.SharePoint.Client.dll")
                        Add-Type -Path $SharePointClientDll 
                        $DownloadURI = New-Object System.Uri($attach.AttachLongPathName);
                         $SharepointHost = "https://" + $DownloadURI.Host
                          $clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($SharepointHost)
                        $soCredentials =  New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($Credentials.UserName.ToString(),$Credentials.password)
                          $clientContext.Credentials = $soCredentials;
                          $FileName = New-UniqueFileName -FileName ($DownloadDirectory + “\” + $attach.Name.ToString())
                          $fileInfo = [Microsoft.SharePoint.Client.File]::OpenBinaryDirect($clientContext, $DownloadURI.LocalPath);
                         $fstream = New-Object System.IO.FileStream($FileName, [System.IO.FileMode]::Create);
                         $fileInfo.Stream.CopyTo($fstream)
                         $fstream.Flush()
                          $fstream.Close()
                         Write-Host ("File downloaded to " + ($FileName))
                    }
                }
            }
        }
    }
    else
    {
        $attachemtItems = New-EWSeDiscoveryPreviewItems -service $service -KQL $KQL -SearchableMailboxString $MailboxName 
        foreach($AttachmentItem in $attachemtItems){
            Write-Host ("Processing Item " + $AttachmentItem.Subject)
            $Item = [Microsoft.Exchange.WebServices.Data.Item]::Bind($service,$AttachmentItem.Id)
            foreach($attach in $Item.Attachments){
                Write-Host $attach
            }
        }
    }
}