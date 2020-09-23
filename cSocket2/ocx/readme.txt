The Winsock Control from Microsoft is not the easiest thing in the world to work with, but it is 
a lot easier than working directly with the Winsock API. Unfortunately, the Winsock Control does 
not support IPv6, and it doesn’t look like it ever will.

So what are the alternatives for VB programmers who do not want to get involved with C++ orC#. 
"Csocket" by Oleg Gdalevich is a drop-in replacement for the Winsock Control. It was further 
enhanced by Emiliano Scavuzzo with his "CsocketMaster" class. There are a number of advantages 
with using a Class Module instead of a Control, not the least of which is that it can be modified 
to suit special needs. These 2 authors took a different approach, but both use "callbacks". VB6 
does not do threading very well, so callbacks are the only real way of communicating with the 
Windows messaging system. The use of callbacks instead of threads is sometimes referred to as 
Non-Blocking calls versus Blocking calls.

If only a few sockets are required, we would recommend using the cSocket2 class and module 
directly. One such program demonstrating this is the included "Chat" program. If however, you need 
multiple sockets for a server type application, then a socket array is the only viable choice. The 
included SMTP Pseudo Server application demonstrates this approach using the cSocket.ocx ActiveX 
Control.

Controls must be registered with the Windows operating system. Visual Basic will do this 
automatically when you compile csocket.ocx. But there are problems with this approach. Every 
time the control is compiled, a new CLSID is created. It is recommended that cSocket.ocx be 
unregistered using regsvr32 before recompiling. Even so, you will probably leave behind multiple 
registry entries. By the time I got a satisfactory product, I had about a dozen excess entries 
in the registry for both the cSocket.ocx and cSocket.oca files. I went through the registry and 
deleted them before the final compilation.

And once you recompile the Control, none of the programs that use it will work anymore. I found 
it easier to use a simple example program (I have included the svcTest program) for the purpose 
of finding the correct CLSID. Delete the old control from the form, add the newly registered 
control to the "Components" of the project, and add the new control back to the form. If you use 
the same name for the object, it will restore any code assigned to it. Verify that it works and 
save the project. Using NotePad or other text editor, copy the object reference line from the .vbp 
file, and paste it over the same line in any other project using the same Control. Likewise, copy 
the object reference line in the .frm file, and paste it over the same line in any other form that 
uses the same Control. In this manner, the other projects should load without complaint from 
Visual Basic, so they can be recompiled.
