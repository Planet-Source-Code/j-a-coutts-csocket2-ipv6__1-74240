SMTP Pseudo Server is a utility program that acts like a Mail Server without actually receiving 
or sending anything. It was originally designed to feed a DNS Black List server, but it has been 
modified to simply receive port 25 SMTP connections and reject them. We found this necessary due 
to continuous bombardment of our DNS server looking for a mail server, and subsequent bombardment 
of our Web server (A record) with port 25 connection attempts. After implementing this program, 
DNS activity is about 1/10 of what it was, and port 25 connection attempts have dropped from over 
2000 to between 600 and 700 per day. The actual port 25 connection attempts to our Web server 
prior to operating the Pseudo Server is undocumented, because port 25 traffic was blocked by our 
router. We only know it was occurring because of Type MX requests followed by Type A requests at
the DNS server.

The SMTP Pseudo Server is designed to run as a service. The need to monitor activity and make 
setting changes in modern Windows environments (Vista and better), necessitated a split 
architecture design to deal with Session Isolation. So there are actually 2 programs; a 
windowless program (SMTPsvc6.exe) to run as a service in session 0, and a form based program 
(SMTPServer6.exe ) to run as a graphic program in a different session.

Microsoft discourages Visual Basic programs being run as a service. The reasons are varied, but 
they all boil down to the fact that services run in Session 0 with system privileges. For that 
reason, care must be execised to ensure that nothing can cause a display to be attempted to the 
GUI. Logging to file is a must instead of using the Msgbox function, and all possible errors 
must be trapped. Untrapped errors that attempt to display will cause an endless loop.

A service program must be able to respond to commands from the Service Manager. This is normally 
done in a separate thread, and VB doesn't handle threads very well. We could put in a timer loop 
that cycles every 300 ms, but that is a bit cludgy. I originally used the Dart Service ActiveX 
Control, but that is rather expensive. So I converted it to use the Microsft NTsvc.ocx Control 
instead. The graphic program however interacts directly with the SCM though API calls.

Another restriction when running as a service is that the HKEY_USERS part of the registry is not 
accessible, so HKEY_LOCAL_MACHINE had to be used instead. The NTsvc.ocx program provides access 
to HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services, but unfortunately that could not be used 
and still allow the graphic part of the program to install the service. So once again we had to 
use direct API calls.

Once started, SMTPServer will show as being "Offline" and must be setup. Make sure the "Show/Log 
Data Events" is checked, and the "Accept Mail" is unchecked. It makes no difference to this 
particular program, but it has the ability to change how the service program works.

1. Click the setup button and the first item is the SMTP Greeting. Again, it doesn't affect how 
this part of the program runs, but it needs something such as "Welcome!".

2. The next thing it wants is the IP address to monitor. The advantage of using TCP/IP to 
communicate between programs is that it can be used remotely as well as local. If you enter a 
local IP address (such as 192.168.0.3), the program allows access to the local service. If a 
foreign address is entered, the program is be be used remotely and there is no access to the 
service. The program supports the loopback addresses 127.0.0.1 and ::1, and port 26 is used to 
monitor the server.

3. Finally, it will want a magic word that allows access. This is done to protect the server from 
unauthorized access.

The SMTPsvc.exe program can be run without installing it as a service, but since it is windowless, 
there is nothing to indicate that it is running. To verify, go to the Command Prompt and enter 
the command "netstat -an". Both ports 25 and 26 should be in the listening mode. Since there is 
no interface, you will have to use the Task Manager to shut down the program. The graphic part of 
the program provides the tools to install and start the service, but it can also be installed and 
uninstalled using the "/install" and "/uninstall" options at the command prompt.

"C:\Program Files\SMTPsvc\SMTPsvc.exe /install"

Once installed as a service, you will probably want to change the Start Type to Automatic, but 
that is your option. Now that the service is running, restart the SMTPServer monitor program. 
This time the program should connect to the service and show its status as being online.

That's rather uninteresting. What we need now is some traffic to monitor. As long as you have the 
Telnet program enabled, you can use it from the command prompt to provide some test traffic.
C:\>telnet 192.168.0.3 25
220 Welcome!06/01/2012 2:41:11 PM -0700
HELO me
250 Hello me from 192.168.0.3, pleased to meet you.
MAIL FROM: anyone@anywhere.com
553 Sender anyone@anywhere.com is Invalid!
Connection to host lost.

The server disconnected because I took too long and it has an inactivity timeout. After about 30 
seconds of no new connections, it closes all connections. The server has been arbitrarily set to 
support 25 simultaneous connections, with the last connection and it's current status shown for 
each socket. That should be more than enough unless there is a problem. One such problem was 
encountered recently with an "AUTH LOGIN" request. Even though the server advertises that it is 
not supported, it did not stop some hacker from attempting many such connections. The server was 
modified to respond that it was an unsupported command rather than just ignoring it, and that 
seems to have solved the problem.

The captured data is also logged to file. Those files are stored in the 
"C:\Program Files\SMTPsvc\logs\" directory by date. You can examine the log files using a text 
editor such as NotePad.

Note: The original IPv4 only version utilizes the Winsock Clone ActiveX Control SocketMaster.ocx 
by Emiliano Scavuzzo. The IPv6 version utilizes my own ActiveX Control called cSocket.ocx, which 
is based on SocketMaster.ocx. The IPv6 version software supports both IPv4 and IPv6, but only 
works on Windows Vista or better systems. The IPv6 version also has the ability to verify the IP 
address entered on setup. This is due to a new Winsock API function available in recent Windows 
operating systems called "getaddrinfo", that was designed to handle the larger IPv6 addresses.
