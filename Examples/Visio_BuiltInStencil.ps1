Import-Module VisioBot3000 -Force

Diagram C:\temp\TestVisio_BuiltinStencil.vsdx
Add-StencilSearchPath "C:\Program Files\Microsoft Office\root\Office16\Visio Content\1031", "D:\OneDrive\Visio\Schablonen"
Stencil Containers -BuiltIn Containers
Stencil Servers -Path SERVER_M.vssx
Shape  WebServer -From Servers -MasterName 'Web Server'

Container Classic -from Containers  


Classic Fred {
        WebServer PrimaryServer
    }