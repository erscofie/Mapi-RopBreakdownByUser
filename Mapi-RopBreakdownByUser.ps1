<#
# Author: Eric Scofield (erscofie@microsoft.com)
# Description:  Extracts individual RopIds for all MapiHttp\Mailbox logs in specified directory, converts each RopId to the associated Rop name, then writes Rop names with ActAsUserEmail to csv.

Script adapted from original ".\RopBreakdownByUser.ps1" which parses RCA logs Matthew Huynh (mahuynh@microsoft.com):
https://blogs.technet.microsoft.com/mahuynh/2014/09/22/enable-rop-logging-in-exchange-2010-and-2013/

# Usage: .\Mapi-RopBreakdownByUser.ps1 "c:\temp\mapihttplogs"
 #>
 
param(
[parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[string]$inputDirectory

)

$ropNames = @(
"Reserved",
"RopRelease",
"RopOpenFolder",
"RopOpenMessage",
"RopGetHierarchyTable",
"RopGetContentsTable",
"RopCreateMessage",
"RopGetPropertiesSpecific",
"RopGetPropertiesAll",
"RopGetPropertiesList",
"RopSetProperties",
"RopDeleteProperties",
"RopSaveChangesMessage",
"RopRemoveAllRecipients",
"RopModifyRecipients",
"RopReadRecipients",
"RopReloadCachedInformation",
"RopSetMessageReadFlag",
"RopSetColumns",
"RopSortTable",
"RopRestrict",
"RopQueryRows",
"RopGetStatus",
"RopQueryPosition",
"RopSeekRow",
"RopSeekRowBookmark",
"RopSeekRowFractional",
"RopCreateBookmark",
"RopCreateFolder",
"RopDeleteFolder",
"RopDeleteMessages",
"RopGetMessageStatus",
"RopSetMessageStatus",
"RopGetAttachmentTable",
"RopOpenAttachment",
"RopCreateAttachment",
"RopDeleteAttachment",
"RopSaveChangesAttachment",
"RopSetReceiveFolder",
"RopGetReceiveFolder",
"Reserved",
"RopRegisterNotification",
"RopNotify",
"RopOpenStream",
"RopReadStream",
"RopWriteStream",
"RopSeekStream",
"RopSetStreamSize",
"RopSetSearchCriteria",
"RopGetSearchCriteria",
"RopSubmitMessage",
"RopMoveCopyMessages",
"RopAbortSubmit",
"RopMoveFolder",
"RopCopyFolder",
"RopQueryColumnsAll",
"RopAbort",
"RopCopyTo",
"RopCopyToStream",
"RopCloneStream",
"Reserved",
"Reserved",
"RopGetPermissionsTable",
"RopGetRulesTable",
"RopModifyPermissions",
"RopModifyRules",
"RopGetOwningServers",
"RopLongTermIdFromId",
"RopIdFromLongTermId",
"RopPublicFolderIsGhosted",
"RopOpenEmbeddedMessage",
"RopSetSpooler",
"RopSpoolerLockMessage",
"RopGetAddressTypes",
"RopTransportSend",
"RopFastTransferSourceCopyMessages",
"RopFastTransferSourceCopyFolder",
"RopFastTransferSourceCopyTo",
"RopFastTransferSourceGetBuffer",
"RopFindRow",
"RopProgress",
"RopTransportNewMail",
"RopGetValidAttachments",
"RopFastTransferDestinationConfigure",
"RopFastTransferDestinationPutBuffer",
"RopGetNamesFromPropertyIds",
"RopGetPropertyIdsFromNames",
"RopUpdateDeferredActionMessages",
"RopEmptyFolder",
"RopExpandRow",
"RopCollapseRow",
"RopLockRegionStream",
"RopUnlockRegionStream",
"RopCommitStream",
"RopGetStreamSize",
"RopQueryNamedProperties",
"RopGetPerUserLongTermIds",
"RopGetPerUserGuid",
"Reserved",
"RopReadPerUserInformation",
"RopWritePerUserInformation",
"Reserved",
"RopSetReadFlags",
"RopCopyProperties",
"RopGetReceiveFolderTable",
"RopFastTransferSourceCopyProperties",
"Reserved",
"RopGetCollapseState",
"RopSetCollapseState",
"RopGetTransportFolder",
"RopPending",
"RopOptionsData",
"RopSynchronizationConfigure",
"Reserved",
"RopSynchronizationImportMessageChange",
"RopSynchronizationImportHierarchyChange",
"RopSynchronizationImportDeletes",
"RopSynchronizationUploadStateStreamBegin",
"RopSynchronizationUploadStateStreamContinue",
"RopSynchronizationUploadStateStreamEnd",
"RopSynchronizationImportMessageMove",
"RopSetPropertiesNoReplicate",
"RopDeletePropertiesNoReplicate",
"RopGetStoreState",
"Reserved",
"Reserved",
"RopSynchronizationOpenCollector",
"RopGetLocalReplicaIds",
"RopSynchronizationImportReadStateChanges",
"RopResetTable",
"RopSynchronizationGetTransferState",
"Reserved",
"Reserved",
"Reserved",
"RopTellVersion",
"Reserved",
"Reserved",
"RopFreeBookmark",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"RopWriteAndCommitStream",
"RopHardDeleteMessages",
"RopHardDeleteMessagesAndSubfolders",
"RopSetLocalReplicaMidsetDeleted",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"RopBackoff",
"Reserved",
"Reserved",
"Reserved",
"Reserved",
"RopLogon",
"RopBufferTooSmall"
)

$inputDirectory = $inputDirectory.ToLower();

if (!$inputDirectory.EndsWith("\*.log")) {
    $inputDirectory += "\*.log"
}

$inputFiles = Get-ChildItem $inputDirectory

write-host "processing directory: $inputDirectory"
#write-host $inputFiles


foreach ($file in $inputFiles) {

    $fileName = $file.FullName
    $input = [System.IO.StreamReader] $fileName
    
    write-host "parsing file" $fileName

    $outputFile = $fileName.ToLower().Replace(".log", "_ropBreakdownByUser.csv")
    $stream = [System.IO.StreamWriter] $outputFile

    $counter = 0

    $lineCount = 0

    $idxUser = 0;
    $idxRopId = 0;

    # write CSV header
    $stream.WriteLine("ActAsUserEmail,Rop")

    while (!$input.EndOfStream) {
        $line = $input.ReadLine()

        # ignore commented lines, but parse the header
        if ($line.StartsWith('#')) {
            # determine indexes of fields: ActAsUserEmail, RopIds
            if ($line.Contains("Fields")) {
                $headers = $line.Replace("#Fields: ", "");
                $headers = $headers.Split(",");
                for ($i = 0; $i -lt $headers.Length; $i++) {
                    $h = $headers[$i];
                    switch ($h) {
                        "ActAsUserEmail" {$idxUser = $i}
                        "RopIds" {$idxRopId = $i;}
                    }
                }
            }
        } else {

            $lineCount++

            #write-host "current line:" $line

            $fields = $line.Split(",");

            $User = $fields[$idxUser]
            #$separator = char[][]
            $rops = $fields[$idxRopId].Split("]")

            foreach ($rop in $rops)
            {
                if ($rop.Contains('>')) {
                    $counter++
                    $newRec = $rop.Trim().Replace('>', '').Replace("[", '').Replace('"', '')        
                    $newRec = $User + "," +  $ropNames[$newRec]
                    #Write-Host $newRec
                    $stream.WriteLine($newRec)
                }
            }
            <#
            if ($lineCount % 10000 -eq 0) {
                write-host "$lineCount lines parsed so far"
            }
            #>

        }
        
    }

    #Write-host "UserIndex: $idxUser, RopIdIndex: $idxRopId"
    #write-host "the last line was" $line

    $input.Close()

    $stream.Flush()
    $stream.Close()

    write-host $counter "ROPs written to" $outputFile
}