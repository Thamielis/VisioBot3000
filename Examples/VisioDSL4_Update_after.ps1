import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio3.vsdx 

# Define shapes, containers, and connectors for the diagram
Stencil Containers -From "C:\GitHub\PowerShell\VisioBot3000\Examples\MyContainers.vssx" 
Stencil Servers -From SERVER_M.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'
Container Logical -From Containers -MasterName 'Logical'
Connector SQL -Color Green -arrow -bidirectional 

#this is the diagram
Logical MyFarm {
    Get-Location MyCity {
        Domain MyDomain_A {
		    WebServer PrimaryServer  
            WebServer SecondaryServer 
	    }
    }

    Get-Location DRSite {
        Domain MyDomain_B {
		    WebServer BackupServer  
  	    }
    }
}
SQL -From PrimaryServer -To BackupServer
Hyperlink $BackupServer -link http://google.com

#Complete-Diagram 