# Mapi-RopBreakdownByUser

Extracts individual RopIds for all MapiHttp\Mailbox logs in specified directory, converts each RopId to the associated Rop name, then writes Rop names with ActAsUserEmail to csv.

Rop logging must be enabled for MapiHttp\mailbox logs so RopId field is populated. Refer to the [Wiki](https://github.com/erscofie/Mapi-RopBreakdownByUser/wiki/Enable-Rop-logging-for-MapiHttp-in-Exchange-2013-and-2016) for more info.

Script adapted from [.\RopBreakdownByUser.ps1](https://blogs.technet.microsoft.com/mahuynh/2014/09/25/rop-breakdown-by-user/), which parses RPC Client Access logs, written by Matthew Huynh (mahuynh<span></span>@microsoft.com).


Usage: .\Mapi-RopBreakdownByUser.ps1 "c:\temp\mapihttplogs"
