# MapiHttp Rop Logging

ROP logging is not enabled by default, mostly because it increases log size significantly and isn't needed to troubleshoot most issues. However, there are times when being able to log individual Rops being sent from the client and those returned from the server is very useful, and even necessary.

## Enable Rop logging for MapiHttp in Exchange 2013/2016:

The process to enable this with MapiHttp is very similar to [enabling ROP logging for RPC Client Access](https://blogs.technet.microsoft.com/mahuynh/2014/09/22/enable-rop-logging-in-exchange-2010-and-2013/). There are enough subtle differences between the two though, so I'll document the steps here to avoid confusion:

To enable ROP logging, you need to edit the web.config file located in the following path:<br>
*C:\Program Files\Microsoft\Exchange Server\V15\ClientAccess\mapi\emsmdb*

This is a per server setting, meaning you must edit the config file on every server you need to enable ROP logging on.

Open web.config in notepad, then find the **LoggingTag**. Add the 'Rops' tag in the value section, as shown here:

<p><img src="img/web.config_enable.png" /></p>

After saving the change, individual Rops will then be included under the 'RopIds' field in MapiHttp mailbox logs, located in the  *"..\V15\Logging\MapiHttp\Mailbox"* directory.

## Making use of the data:

For those that have used Rop logging with RPC Client Access logs before, you'll recall that Rops are logged using the Rop names, making it easy to decipher. With MapiHttp, RopIds are logged rather than the name, so you end up with something that looks like this (highlighting the ROPs in from client, indicated by the '>' character):


