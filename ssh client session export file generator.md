# ssh client session export file generator

When you need dynamic changed ssh client session for large scale of networking, or need to migrate session between different workstation, one click to create a export session file will be handy.

SecureCRT and mobaXterm are most common used ssh clients in daily work.

## How to run it
- build up host access info table in excel: column Hostname,HostIP,RemotePort,Username
- run macro from excel Developer tab or create a customzied ribbon
  
## VBA macro list 

**Sub ScrtExport()**

secureCRT export session xml generator
- input data is selection of column Hostname,HostIP,RemotePort,Username
- will generate session file .\Export\Session\scrt-\<active-sheet\>-\<timestamp\>.xml
- all sessions will be under folder which name from active sheet
- open generated secureCRT session file in notepad for review

**Sub mobaExport()**

mobaXterm export session mxtsessions generator
- input data is selection of column Hostname,HostIP,RemotePort,Username
- will generate session file .\Export\Session\\mobaxterm-\<active-sheet\>-\<timestamp\>.mxtsessions
- all sessions will be under folder which name from active sheet
- open generated mobaXterm session file in notepad for review
