ipmo VisioBot3000 -Force

Diagram C:\temp\TestVisio_BuiltinStencil.vsdx 
Stencil Containers -BuiltIn Containers
Stencil Servers -Path C:\temp\SERVER_U.vssx
Shape  WebServer -From Servers -MasterName 'Web Server'

Container Classic -from Containers -MasterName Classic


Classic Fred {
        WebServer PrimaryServer
    }