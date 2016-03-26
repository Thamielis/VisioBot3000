Import-Module VisioBot3000 -Force

New-VisioApplication
New-VisioDocument c:\temp\TestVisio.vsdx 
Register-VisioStencil -Name Servers -path 'C:\Program Files (x86)\Microsoft Office\Office15\Visio Content\1033\SERVER_U.VSSX‘
Register-VisioShape -name WebServer -StencilName Servers -masterName 'Web Server'
New-VisioShape WebServer -x 5 -y 5
