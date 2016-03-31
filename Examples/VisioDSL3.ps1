stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio3.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio3.vsdx 
Stencil Containers -From C:\temp\MyContainers.vssx 
Stencil Servers -From C:\temp\SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'
Container Logical -From Containers -MasterName 'Logical'
Connector SQL -Color Red -arrow


Logical MyFarm {
    Location MyCity {
        Domain MyDomain {
		    WebServer PrimaryServer 5 5
	    }
    }
    Location DRSite {
        Domain MyDomain {
		    WebServer BackupServer 5 8
	    }
    }
    SQL -From PrimaryServer -To BackupServer
}