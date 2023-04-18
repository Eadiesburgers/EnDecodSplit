# EnDecodSplit
VBScript tool to base64 encode and split a large file of any type into multiple configurable smaller files, or join and decode a bunch of files.
Additionally, assists in circumventing banned email attachment file type policy by temporarily corrupting the file header metadata of the first file, to ensure that unless all par files received are together and the same script is used to decode, email relay gateways or email servers will be unaware of the file type in transit.

You have a large file you need to email but it far exceeds your organisation email attachment size policy.
Additionally, you have a file type that needs sending / receiving that is banned by your organisations email attachment policy.
It would take longer to provision an exemption case and there are no COTS solutions that can resolve this issue in your APL.
Here's a VBScript utility to the rescue, baked in Notepad.exe.

The utility base64 encodes and splits-up a larger file into file sizes of your choice, with no two files being of the same size.  This assists the email attachments getting to their intended estination, should any old school email IDS technologies be in place.
The utility also performs the reverse, in joining the split files and base64 decoding them, back to their original single file.

Why VBScript? It'll run on any Windows platform post 1996, right up to the latest (to date) Windows 11 edition, including Insider pre-release builds.  At some point in time it would have been pivotal in the support of your IT estate and this enduring backward compatibility and fear of decommissioning activities breaking something ensures it lives on.  It's easy to understand and thus own and extend with features.  Better still it's difficult to lock-down at a granular level, unlike PowerShell with it's tunable execution policy.  Therefore if cscript/wscript binaries work, they'll work in their entirety, as the dependency files are few and far between. 

Run via console cscript binary (cscript.exe) or Win app binary (wscript.exe).  Pass no parameters for help text guidance to be displayed.

cscript /nologo endecodsplit.vbs


