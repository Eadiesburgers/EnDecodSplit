# EnDecodSplit
VBScript tool to base64 encode and split a large file of any type into multiple configurable smaller files or join and decode a bunch of files.

You have a large file you need to email but it far exceeds your companies email attachment size policy.
It would take longer to provision an exemption case and there are no COTS solutions that can resolve this issue in your APL.
Then here's a VBScript utility to the rescue, baked in Notepad.exe.

The utility base64 encodes and splits-up a larger file into file sizes of your choice, with no two files being of the same size.  This assists the email attachments getting to their intended estination, should any old school email IDS technologies be in place.
The utility also performs the reverse, in joining the split files and base64 decoding them, back to their original single file.

Why VBScript? It'll run on any Windows platform post 1996, right up to the latest (to date) Windows 11 edition, including Insider pre-release builds.  At some point in time it would have been pivotal in the support of your IT estate and this enduring backward compatibility and fear of decommissioning activities breaking something ensures it lives on.  It's easy to understand and thus own and extend with features.  Better still it's difficult to lock-down at a granular level, unlike PowerShell with it's tunable execution policy.  Therefore if cscript/wscript binaries work, they'll work in their entirety, as the dependency files are few and far between. 

Run via console cscript binary (cscript.exe) or Win app binary (wscript.exe):  cscript /nologo endecodesplit.vbs
Pass no parameters for help text.
