Telned


Introduction

When all the little trojans started floating about the internet it
seemed like an interesting idea so I decided to make my very own.
I havent really got a name for it yet and I am definately not planning 
to release it. It is fully functional supporting full file system
manipulation including del,mkdir,rmdir,shell,upload,download,dir, 
echofile, drive, etc...
The program can also list running tasks, retrieve full user information
(including timezone information) and create datapipes on specified ports
(so you can use the server system as a single connection proxy for example)

About the code

The code is very very messy. I made up functions as I went along.
Some of the code such as the upload and download code is very raw and
requires a little more work perhaps.

Except for that its a very nicely functioning program. A random port
in the temporary port range is picked every time the server is started,
full error handling and remote reporting is supported.

I intended it to be a personal single user system so I did not index
the main TCP sockets to allow for multiple connections. 

Known Bugs:

Download often fails to report that its stop sending data. Im not
sure why this but winsock refuses if the lines after close the 
connection. None the less the function still works just fine you just
have to decide when to stop downloading.

No stealth code implemented yet even though I've already written a
function to allow the server to remain resident on the remote system.
If you wish to use stealth code you may find lots and lots of stealth
examples at www.planet-source-code.com/vb



WARNING: Due to the nature of this code installation of its program
files on a system without permission can be considered invasion of
privacy and is a prosecutable offence so please do not use this for
any criminal puposes. It is here so you can learn from it.

Comments,questions email Emir.U -> unknown@foreigner.co.uk