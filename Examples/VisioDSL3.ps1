#clean-up because I'm re-running this over and over
stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio3.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio3.vsdx 

# Define shapes, containers, and connectors for the diagram
Stencil Containers -From C:\temp\MyContainers.vssx 
Stencil Servers -From C:\temp\SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'
Container Logical -From Containers -MasterName 'Logical'
Connector SQL -Color Red -arrow 

#this is the diagram
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
}
SQL -From PrimaryServer -To BackupServer
