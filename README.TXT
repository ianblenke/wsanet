                              WSANet.VBX v1.08a
                Windows Sockets Networking Visual Basic Control

                             December 21, 1993

                             Author: Ian Blenke
                        VB Suggestions: G. Michael Carr

              THIS IS AN ALPHA RELEASE. PLEASE KEEP DISTRIBUTIONS
                         IN COMPETENT HANDS ONLY.


Purpose:

    To encapsulate Windows Sockets v1.1 C calls and aid in
quick network application development by adding an extention control to
Visual Basic.


Status:

-v1.00a:
-    This is an alpha release. It is buggy. It will remain buggy until I
-find the time and inclination to track them all down. Feel free to help
-me in this venture, as it is likely to continue through to the release
-version.

-v1.01a:
-    Patch added to allow setting of LocalPort property. If set to 0,
-an arbitrary port number will be assigned by your WinSock app.
-    Now release 2 version of the VBX: WSMTPC.VBX and RELEASE.VBX. Both
-are the same source code, but the RELEASE.VBX does not have debugging
-information either in OutputDebugString() form or by compiler symbols.
-In order to use this smaller "release" version, just rename WSNETC.VBX
-to DBWSNETC.VBX (or equivalent) and copy RELEASE.VBX to WSNETC.VBX.

-v1.02a
-    Internal minor overhaul of code
-    Properties have changed "order", so this version will no longer work
-with EXEs meant for previous versions.
-    Not released due to bug fixes

-v1.03a
-    Context sensitive help added
-    Properties added: LocalService/RemoteService and SendThreshold
-    OnSend() semantics have changed
-    Added NetClientSendBlock() and NetClientRecvBlock() functions - BUT THEY
- HAVEN'T BEEN TESTED YET.
-    The Multiple events bug has been fixed - posting VBM_FIREEVENT messages
- was a bad thing to do.
-    Many minor bugs fixed.

-v1.04a
-    Context sensitive completed. It's now awaiting criticism.
-    OnTimeOut() semantics are now documented
-    A Few dangerous bugs are now fixed
-    VC++ should compile with fewer errors now, although the errors
      "C:\WSNETC\WSNetC.H(3) : warning C4001: nonstandard extension used -
       'single line comment'" and
      "c:\wsnetc\init.c(26) : warning C4305: 'cast' : truncation from
       'unsigned int __far *' to 'unsigned short '"
     will persist until I find a #pragma or trick to silence them.
-    VC++ Makefile added to the distribution (thanks Dan!)
-    Possible problems with WinRSH still exist (like, it doesn't work)

-v1.05alpha
-    Socket semantics have changed (although the ability to set it is not
      yet technically implemented).
-    New Toolbox BMP has been designed
-    HLP updated: Changes topic added.
-    The LocalPort property was broken previous to version 1.05alpha. Setting
      it would set the RemotePort property.
-    WSNETC.C and NETC.C now use the DEBUG_BUILD flag to mark debug/release
      build versions (instead of keeping two copies of source).
-    WSNETC.C cleaned up. Debugging information clearer now.
-    NETC.C cleaned up. Debugging information clearer now.
-    NET.C - Added "netConnect()" procedure to aid in the setting of the
      Socket property
-    All sources cleaned up for compilation by VC++ (uncomment USE_VC++ in
      WSNetC.H)
-    Fixed TimeOut bug. Previous versions will ALWAYS delay for just 1 second.

-v1.06alpha
-    Added Version and Debug properties.
-    Added HostAliasCount/HostAliasList() and
      HostAddressCount/HostAddressList() properties
-    Changed the HostAddr/HostName property code to avoid VB calls where
      possible.
-    Added Ini Control - 6 properties, no events
-    Changed the names of EVERYTHING
-    Moved everything around
-    Caused general havok within the code:
      Init.C is renamed to WSANet.C, and 2 procedures from WSNetC.C were
       added.
      WSNetC.C is copied over to NetClnt.C
      Ini.C/Ini.H is added (to support the Ini control)
      WSNetC.H becomes two files: NetClnt.H and WSANet.H
      Functions added to Net.C: GetAtomHsz(), SetAtom()
     Reason for changes:
      NetClnt.C/NetClnt.H now hold all of the NetClient control code
      Ini.C/Ini.H now hold all of the Ini control code
      WSANet.C/WSANet.H is the "glue" for the project.
      Version.H holds all version information
      WSANet.RC/WSANet.DEF hold the resources and definitions
-    Caused the programmer many headaches as he renamed everything ;)
-    Wrote a new example client: WiNslookup
-    Added 2 sample clients to the HLP file.

-v1.07alpha
-    Added the NetServer control.
-    Fixed a nasty problem with v1.06alpha's use of ATOMs to store strings:
      all strings are now stored internally as Visual Basic strings (type
      HLSTR).
-    Changed the semantics of setting the Socket on the NetClient control so
      that NetServer's OnAccept(Socket, ...) could be handed over.
-    Fixed a potential problem with the FD_CLOSE message handler: it didn't
      "zero out" the Socket when it closed!
-    Dropped the VBTest client due to it's obvious inadequacies (and
      pointlessness)
-    Wrote WinWhois client.
-    Wrote WFingerD server.
-    Modified all VB3.0 examples to use a "standard" GUI appearance.
-    Shuffled the HLP file around. Topics have changed to include the control
      name along with the property context string.

-v1.08alpha (you are reading a pre-alpha version of this release)
-    Fixed a problem that caused WSAE_NOCONN errors.
-    Fixed a event firing problem.
-    Fixed the connection semantics
-    Added window position "memory" to VB3.0 samples
-    Wrote VB-SMTPD server. Multiple incoming connections!
-    Added LISTEN_DONTREUSE feature for the Listen property.
-    Added the Host property
-    Fixed a problem with the Line property (old versions can GPF randomly)

To-Do list:

- Make sure the binary transfer properties and functions are working
    correctly (they are known to be broken).
- Make VC++ "sample" applications (might be possible for me to do this
    soon)
- Add OnGetHost() event, and AsyncHandle property (for non-blocking DNS
    lookups)
- Add TCP/UDP and OOB switch.


Requirements:

(Run time, Design time)
    A TCP/IP stack with their WinSock.DLL v1.1
    VB1.0+

(When recompiliing the control)
    The CDK that comes with the professional versions of VB1.0/VB2.0/VB3.0 or
     VC++
    The only "extra" library needed is WINSOCK.LIB (which is provided).


Copyright:

In this document, VBX refers to the actual Visual Basic control distributed
with and based on this WSNETC project, and author refers to Ian Blenke. The
author cannot be held responsible for damages, expressed or implied, for the
use of this software. No commercial use can be made of this product without
the consent of the author. No profit of any kind can be made on the sale or
distribution of this program. If you wish to distribute this program with
other samples of WinSock programming, you must first contact the author so
that he can keep accurate records of its usage; no charge may be made for
this project's availability other than the cost of the physical medium used
to store it on and any processing costs. If you write any programs based on
this source code, including parts of this source code, or in some way derived
from this source code, you may not sell them for any profit without the
written consent of the author. If you incorporate this VBX into a public
domain program, all the author requires is a notification that "The NetClient
control was written by Ian Blenke" and some form of notification that his name
was used in the public domain software distribution. No profit of any kind,
shareware/commercial or otherwise, may be made for software based on this VBX
without the written consent of the author. This does not represent a contract
on the part of the author. If any issues cannot clearly be resolved by reading
this text, immediately contact the author.

The VBX control distributed with this project must be distributed with your
programs free of charge. You cannot charge money for any program based on the
VBX without the written consent of the author.

Notes:

If you have any bug reports and/or source patches... By all means, tell me! I
would be glad to help keep this code up to date. Do not, however, modify this
source in any way and re-distribute it without the author's knowledge. Follow
the above Copyright, and all will be well.

I don't like such agreements, but in today's world of lawyers and lawbreakers,
I have little other choice. Enjoy!

(This is meant to scare code "scalpers". Those that are actually learning C
 programming for Visual Basic controls and WinSock applications are free to
 adapt any of the ideas of this code. Remember, new code should be better
 code, so "adapt" this, do NOT copy it.)


Please consult the .HLP file for furthur information.


