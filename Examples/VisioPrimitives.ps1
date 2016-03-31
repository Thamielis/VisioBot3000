Import-Module VisioBot3000 -Force

New-VisioApplication

New-VisioDocument C:\temp\TestVisio3.vsdx 
Register-VisioStencil -Name Containers -Path C:\temp\MyContainers.vssx 
Register-VisioStencil -Name Servers -Path C:\temp\SERVER_U.vssx
Register-VisioShape -Name WebServer -From Servers -MasterName 'Web Server'
Register-VisioContainer -Name Location -From Containers -MasterName 'Location'
Register-VisioContainer -Name Domain -From Containers -MasterName 'Domain'
Register-VisioContainer -Name Logical -From Containers -MasterName 'Logical'

New-VisioContainer -shape (Get-VisioShape Logical) -label MyFarm -contents {
    New-VisioContainer -shape (Get-VisioShape Location) -label MyCity -contents {
        New-VisioContainer -shape (Get-VisioShape Domain) -label MyDomain -contents {
		    New-VisioShape -master WebServer -label PrimaryServer -x 5 -y 5
	    }
    }
    New-VisioContainer -shape (Get-VisioShape Location) -label DRSite -contents {
        New-VisioContainer -shape (Get-VisioShape Domain) -label MyDomain -contents {
		    New-VisioShape -master WebServer -label BackupServer -x 5 -y 8
	    }
    }
    New-VisioConnector -From PrimaryServer -To BackupServer -Label SQL -Color Red -Arrow
}