# JunkMailStop
Outlook addin to filter junk mail by parsing the mail headers
Outlook addin that parses the mail headers and looks for servers that are known to send junk mail. The user can add and remove servers from a black list.

Currently the only parsing that the addin is doing is extracting the servers that the mail was received from and comparing them to black servers list.
The addin adds a "Servers" user property to each received mail item and shows it in a new column named "Servers". The column is shown in the outlook view pane.
The user can right click a mail item or multiple mail items and choose mail servers for adding to the black servers list. The list of mail servers that are shown to the user contains mail servers extracted from the headers of the selected mail item(s).
The addin also supports a configuration property page shown in outlook's addins options dialog.
The entire code is C++ with ATL objects.
This is a Visual Studio 2010 project. The addin was developed for outlook 2010, wasn't tested on other versions of outlook.
The project build jmsoaddin.dll (Junk Mail Stop Outlook Addin). To install the addin just run:
regsvr32 jmsoaddin.dll
I use TortoiseSVN for my development. There is a pre-build step that sets the revision number according to the SVN revision. You will have to work this around.
Also in stdafx.h there are 3 #import lines with absolute path to office files. You'll have to modify this according to your system.
