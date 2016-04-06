import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio3.vsdx -Update  

# Define shapes, containers, and connectors for the diagram
Stencil Containers -From C:\temp\MyContainers.vssx 
Stencil Servers -From C:\temp\SERVER_U.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'
Container Logical -From Containers -MasterName 'Logical'
Connector SQL -Color Green -arrow -bidirectional 

#this is the diagram
Logical MyFarm {
    Location MyCity {
        Domain MyDomain_A {
		    WebServer PrimaryServer 50 50
            WebServer SecondaryServer 7 3
	    }
    }

    Location DRSite {
        Domain MyDomain_B {
		    WebServer BackupServer 50 80
	    }
    }
}
SQL -From PrimaryServer -To BackupServer
Hyperlink $BackupServer -link http://google.com

#Complete-Diagram 