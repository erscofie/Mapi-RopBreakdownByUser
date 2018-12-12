# Mapi-RopBreakdownByUser
Usage: `.\Mapi-RopBreakdownByUser.ps1 "c:\temp\mapihttplogs"`

Extracts individual RopIds for all MapiHttp\Mailbox logs in the specified directory, converts each RopId to the associated Rop name, then writes the Rop name along with user and client info for each rop to CSV. 

*One .CSV file is created for each .LOG file processed.*

**Rop logging must be enabled for MapiHttp\mailbox logs so RopId field is populated.** Refer to the [Wiki](https://github.com/erscofie/Mapi-RopBreakdownByUser/wiki/Enable-Rop-logging-for-MapiHttp-in-Exchange-2013-and-2016) for more info.

Script adapted from [.\RopBreakdownByUser.ps1](https://blogs.technet.microsoft.com/mahuynh/2014/09/25/rop-breakdown-by-user/), which parses RPC Client Access logs, written by [Matthew Huynh](https://github.com/maxxwizard).

Thanks to [Bill Long](https://github.com/bill-long) for contributing guidance and recommendations for improvements
