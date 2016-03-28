Diagram C:\temp\TestVisio.vsdx -from c:\temp\Template.vstx
Stencil Containers -BuiltIn
Stencil Servers -From C:\stencil.vssx
Shape WebServer –From Servers –Name ‘Web Server’
Shape Location –From Containers –Name 'Physical Location'
Shape Domain -From Containers -Name 'AD Domain'
Domain MyDomain {
	Location MyCity {
		WebServer PrimaryServer
	}
	Location DRSite {
		WebServer BackupServer
	}
}
