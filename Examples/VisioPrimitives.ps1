Import-Module VisioBot3000 -Force

New-VisioApplication

New-VisioDocument C:\temp\VisioPrimitives1.vsdx 
Register-VisioStencil -Name Containers -Path C:\temp\MyContainers.vssx 
Register-VisioStencil -Name Servers -Path C:\temp\SERVER_U.vssx
Register-VisioShape -Name WebServer -From Servers -MasterName 'Web Server'
Register-VisioContainer -Name Location -From Containers -MasterName 'Location'
Register-VisioContainer -Name Domain -From Containers -MasterName 'Domain'
Register-VisioContainer -Name Logical -From Containers -MasterName 'Logical'

New-VisioContainer -shape (Get-VisioShape Logical) -Name MyFarm -contents {
    New-VisioContainer -shape (Get-VisioShape Location) -Name MyCity -contents {
        New-VisioContainer -shape (Get-VisioShape Domain) -Name MyDomain -contents {
		    New-VisioShape -master WebServer -Name PrimaryServer -x 5 -y 5
	    }
    }
    New-VisioContainer -shape (Get-VisioShape Location) -Name DRSite -contents {
        New-VisioContainer -shape (Get-VisioShape Domain) -Name MyDomain -contents {
		    New-VisioShape -master WebServer -Name BackupServer -x 5 -y 8
	    }
    }
    New-VisioConnector -From PrimaryServer -To BackupServer -Name SQL -Color Red -Arrow
}