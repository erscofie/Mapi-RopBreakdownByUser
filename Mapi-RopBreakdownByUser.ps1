<#####
# v1.01
# Author: Eric Scofield (erscofie@microsoft.com)
# Description:  Extracts individual RopIds for all MapiHttp\Mailbox logs in specified directory, converts each RopId to the associated Rop name, then writes Rop names with ActAsUserEmail to csv.
#
# Script adapted from original ".\RopBreakdownByUser.ps1" which parses RCA logs Matthew Huynh (mahuynh@microsoft.com):
# https://blogs.technet.microsoft.com/mahuynh/2014/09/22/enable-rop-logging-in-exchange-2010-and-2013/
#
# Usage: .\Mapi-RopBreakdownByUser.ps1 "c:\temp\mapihttplogs"
#####>

param(
[parameter(Mandatory=$true)]
[ValidateNotNullOrEmpty()]
[string]$inputDirectory

)

$ropNames = @{
1 = "RopRelease";
2 = "RopOpenFolder";
3 = "RopOpenMessage";
4 = "RopGetHierarchyTable";
5 = "RopGetContentsTable";
6 = "RopCreateMessage";
7 = "RopGetPropertiesSpecific";
8 = "RopGetPropertiesAll";
9 = "RopGetPropertiesList";
10 = "RopSetProperties";
11 = "RopDeleteProperties";
12 = "RopSaveChangesMessage";
13 = "RopRemoveAllRecipients";
14 = "RopModifyRecipients";
15 = "RopReadRecipients";
16 = "RopReloadCachedInformation";
17 = "RopSetMessageReadFlag";
18 = "RopSetColumns";
19 = "RopSortTable";
20 = "RopRestrict";
21 = "RopQueryRows";
22 = "RopGetStatus";
23 = "RopQueryPosition";
24 = "RopSeekRow";
25 = "RopSeekRowBookmark";
26 = "RopSeekRowFractional";
27 = "RopCreateBookmark";
28 = "RopCreateFolder";
29 = "RopDeleteFolder";
30 = "RopDeleteMessages";
31 = "RopGetMessageStatus";
32 = "RopSetMessageStatus";
33 = "RopGetAttachmentTable";
34 = "RopOpenAttachment";
35 = "RopCreateAttachment";
36 = "RopDeleteAttachment";
37 = "RopSaveChangesAttachment";
38 = "RopSetReceiveFolder";
39 = "RopGetReceiveFolder";
41 = "RopRegisterNotification";
42 = "RopNotify";
43 = "RopOpenStream";
44 = "RopReadStream";
45 = "RopWriteStream";
46 = "RopSeekStream";
47 = "RopSetStreamSize";
48 = "RopSetSearchCriteria";
49 = "RopGetSearchCriteria";
50 = "RopSubmitMessage";
51 = "RopMoveCopyMessages";
52 = "RopAbortSubmit";
53 = "RopMoveFolder";
54 = "RopCopyFolder";
55 = "RopQueryColumnsAll";
56 = "RopAbort";
57 = "RopCopyTo";
58 = "RopCopyToStream";
59 = "RopCloneStream";
62 = "RopGetPermissionsTable";
63 = "RopGetRulesTable";
64 = "RopModifyPermissions";
65 = "RopModifyRules";
66 = "RopGetOwningServers";
67 = "RopLongTermIdFromId";
68 = "RopIdFromLongTermId";
69 = "RopPublicFolderIsGhosted";
70 = "RopOpenEmbeddedMessage";
71 = "RopSetSpooler";
72 = "RopSpoolerLockMessage";
73 = "RopGetAddressTypes";
74 = "RopTransportSend";
75 = "RopFastTransferSourceCopyMessages";
76 = "RopFastTransferSourceCopyFolder";
77 = "RopFastTransferSourceCopyTo";
78 = "RopFastTransferSourceGetBuffer";
79 = "RopFindRow";
80 = "RopProgress";
81 = "RopTransportNewMail";
82 = "RopGetValidAttachments";
83 = "RopFastTransferDestinationConfigure";
84 = "RopFastTransferDestinationPutBuffer";
85 = "RopGetNamesFromPropertyIds";
86 = "RopGetPropertyIdsFromNames";
87 = "RopUpdateDeferredActionMessages";
88 = "RopEmptyFolder";
89 = "RopExpandRow";
90 = "RopCollapseRow";
91 = "RopLockRegionStream";
92 = "RopUnlockRegionStream";
93 = "RopCommitStream";
94 = "RopGetStreamSize";
95 = "RopQueryNamedProperties";
96 = "RopGetPerUserLongTermIds";
97 = "RopGetPerUserGuid";
99 = "RopReadPerUserInformation";
100 = "RopWritePerUserInformation";
102 = "RopSetReadFlags";
103 = "RopCopyProperties";
104 = "RopGetReceiveFolderTable";
105 = "RopFastTransferSourceCopyProperties";
107 = "RopGetCollapseState";
108 = "RopSetCollapseState";
109 = "RopGetTransportFolder";
110 = "RopPending";
111 = "RopOptionsData";
112 = "RopSynchronizationConfigure";
114 = "RopSynchronizationImportMessageChange";
115 = "RopSynchronizationImportHierarchyChange";
116 = "RopSynchronizationImportDeletes";
117 = "RopSynchronizationUploadStateStreamBegin";
118 = "RopSynchronizationUploadStateStreamContinue";
119 = "RopSynchronizationUploadStateStreamEnd";
120 = "RopSynchronizationImportMessageMove";
121 = "RopSetPropertiesNoReplicate";
122 = "RopDeletePropertiesNoReplicate";
123 = "RopGetStoreState";
126 = "RopSynchronizationOpenCollector";
127 = "RopGetLocalReplicaIds";
128 = "RopSynchronizationImportReadStateChanges";
129 = "RopResetTable";
130 = "RopSynchronizationGetTransferState";
134 = "RopTellVersion";
137 = "RopFreeBookmark";
144 = "RopWriteAndCommitStream";
145 = "RopHardDeleteMessages";
146 = "RopHardDeleteMessagesAndSubfolders";
147 = "RopSetLocalReplicaMidsetDeleted";
249 = "RopBackoff";
254 = "RopLogon";
255 = "RopBufferTooSmall"
}

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

    $idxDateTime = 0;
    $idxUser = 0;
    $idxClient = 0;
    $idxClientVersion = 0;
    $idxRopId = 0;

    # write CSV header
    $stream.WriteLine("DateTime,ActAsUserEmail,Client,ClientVersion,Rop")

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
                        "DateTime" {$idxDateTime = $i}
                        "ActAsUserEmail" {$idxUser = $i}
                        "ClientSoftware" {$idxClient = $i}
                        "ClientSoftwareVersion" {$idxClientVersion = $i}
                        "RopIds" {$idxRopId = $i;}
                    }
                }
            }
        } else {

            $lineCount++

            #write-host "current line:" $line

            $fields = $line.Split(",");
            $DateTime = $fields[$idxDateTime]
            $User = $fields[$idxUser]
            $Client = $fields[$idxClient]
            $clientVersion = $fields[$idxClientVersion]
            #$separator = char[][]
            $rops = $fields[$idxRopId].Split("]")

            foreach ($rop in $rops)
            {
                if ($rop.Contains('>')) {
                    $counter++
                    $newRec = $rop.Trim().Replace('>', '').Replace("[", '').Replace('"', '')
                    $newRec = $DateTime + "," + $User + "," + $Client + "," + $clientVersion + "," +  $ropNames[[int]$newRec]
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
