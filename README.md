# Mapi-RopBreakdownByUser

Extracts individual RopIds for all MapiHttp\Mailbox logs in specified directory, converts each RopId to the associated Rop name, then writes Rop names with ActAsUserEmail to csv.

Rop logging must be enabled for MapiHttp\mailbox logs so RopId field is populated.

Script adapted from original ".\RopBreakdownByUser.ps1" which parses RCA logs, written by Matthew Huynh (mahuynh@microsoft.com):
https://blogs.technet.microsoft.com/mahuynh/2014/09/22/enable-rop-logging-in-exchange-2010-and-2013/

Usage: .\Mapi-RopBreakdownByUser.ps1 "c:\temp\mapihttplogs"
