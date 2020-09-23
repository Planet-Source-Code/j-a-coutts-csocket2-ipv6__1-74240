IPv4 uses a 32 Bit (4 Byte) address that can be dealt with as long integer. At 128 Bits, the IPv6 
address cannot. It must be dealt with as a string or byte array. Many of the IPv4 calls use the 
"sockaddr_in" structure (16 bytes), whereas the "sockaddr_in6" structure is 28 bytes in length. 
To facilitate a common structure, we will use the "sockaddr" structure at a total of 28 bytes. 
This simplified structure only defines "family" and "data", and must be copied into the other 
structures as necessary to extract specific information. Many of the Winsock2 calls had to be 
redefined to accept the "sockaddr" structure. In many ways, the IPv6 calls are much simpler than 
their predecessors, but the sequence of events had to be changed to take advantage of this 
simplicity. The single biggest change is the introduction of the "getaddrinfo()" call, which 
replaces the "gethostbyname()" and "gethostbyaddr()" calls. You simply supply the domain server 
name, and it returns the "IPv4" or "IPv6" address all set up within the "sockaddr" structure. 
To handle the longer IP address, we must use the "inet_ntop()" call instead of "inet_ntoa()".

The "CHAT" program included with the "cSocket2" class and it’s helper module "mWinsock2", 
functions as both a server and client in both IPv4 and IPv6 modes. "CHAT" will listen on either 
IPv4 or IPv6, but not both. This is due to the fact that the Winsock Control was designed to 
control a single socket. If you use the command "ipconfig /all", and you have both IP stacks 
operational, you will notice the same ports listening on "0.0.0.0" and "::". These are the  
"INADDR_ANY"  and "INADDR_ANY6" addresses. We can use these addresses to listen on any available 
addresses, but that still leaves us with 2 distinct sockets. Since our control/class can only 
control a single socket, the user must choose between IPv4 and IPv6

There can also be more than one IP address for the protocol selected, so we can choose from the 
available addresses. If you click on address resolution (GetAddress) for the domain server 
selected, the program will return all available IP addresses for the version selected, from 
which you will be able to select the one to use. If there is only one, it will automatically 
select the only one available. If you connect directly to the server by name, it will use the 
first one returned for the IP protocol selected, which may or may not be the one desired.

When operating in the development environment, there are numerous debug messages to assist the 
developer. Since some problems occur only in the compiled program, there is an additional debug 
facility available. If the program is started with the argument "debug", the program will log 
these same events to a logfile called "IPv6Chat.Log".

The "CHAT" program operates in UDP mode as well as TCP mode. TCP establishes a connection with 
the remote host before transmission is commenced, but UDP is connectionless. Data is simply sent 
to a destination, and the sender has no idea whether it made it to the destination or not. In 
"UDP", the "CHAT" program uses the same commands as the "TCP" mode, but some of them have limited 
functionality. In "UDP" mode, the "Listen" command simply binds the socket to the local address 
("INADDR_ANY" or "INADDR_ANY6"). The "Connect" command also binds to the local socket, but more 
importantly it enables the "Send" button. A connect is not essential, and a person could just as 
easily send directly to a destination without it, but because we are using a common interface, 
it must be utilised.

Switching between types and protocols after starting a communication is not guaranteed to work at 
this point in time. There are simply too many different scenarios that can be encountered to be 
reliable. If you have difficulty at any point, exit the program and restart it. I could have 
offered different programs for each type and protocol in both client and server modes, but that 
would have involved a substantially greater number of programs.

When using "cSocket2", there are some rules. First examine the Fmain form load procedure.

·  The very first routine loads the "cSocket2" class. This routine is essential. 

·  The second routine is the Debug logging routine. It is not essential, but if you need to 
troubleshoot a program later on, it can come in handy.

·  The next routine establishes the session type (TCP or UDP). Since TCP is zero and UDP is one, 
it will default to TCP if not defined.

·  Next comes a routine to establish the IP protocol to be used. I could have used 0 or 1 as in 
the session type, but there are a number of other possible settings that can be used and I 
arbitrarily chose 4 to represent IPv4 and 6 to represent IPv6. The setting of the IP protocol is 
essential for "cSocket2" to work properly.

·  Next comes a routine to find and save all the local IP addresses. This is not an essential 
routine at this point in time.

If a program written with "cSocket2" is attempted to be run on a Windows computer that does not 
support the new calls, it will return an entry point error when the "ws2_32.dll" is accessed. A 
cleaner way of handling the issue would be to check the OS version and warn the operator.

The Status bar on the bottom is run off a timer that I arbitrarily set to 1 second (1000 ms). The 
Winsock Control is limited to controlling a single socket, and the current state of that socket is 
returned in the "State" property of "cSocket2". For convenience, I added a routine that converts 
that state number to a description. When an error is encountered, that error is written to the 
status bar and the timer disabled. Using the "Close" button attempts to clear the error and 
re-enable the timer.

All routines that required a substantial rewrite to work in IPv6 have had a "2" added to the name. 
This so that I could leave the original routine in place while I worked on the new routine. Once 
operational, the old routine was removed. You may also notice several Function Declarations with 
a 2 added, but they alias back to the original function. Visual Basic complains when the variables 
or structures used are not the same as the declarations.

The primary reason for the extra "Debug Logging" routine was to troubleshoot a couple of problems 
that occurred only in the compiled program and not in the IDE. I had to modify some of the Function 
Declarations to support multiple variable types. This is done through the use of the "As Any" 
keyword. In both cases, the problem turned out to be in the "CopyMemory" function. This function 
apparently has some peculiarities which must be observed. When the source or the destination is a 
Visual Basic variable, it should be passed by reference. And when the source or the destination is 
a memory location you should pass it by value. Since the function is declared "As Any": 
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, source As Any, ByVal 
bytes As Long)
That means that it must be passed in the call "ByVal":
CopyMemory ByVal address, lngValue, 4
The default value is "By Ref", so only the "By Val" is required. Additionally, when the source or 
the destination is an array of numbers, you must pass the first element of the array by reference.
