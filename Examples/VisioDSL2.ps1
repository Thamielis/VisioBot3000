stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio2.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force
Diagram C:\temp\TestVisio2.vsdx 
Stencil Containers -From "C:\GitHub\PowerShell\VisioBot3000\Examples\MyContainers.vssx" 
Stencil Servers -From SERVER_M.vssx
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'
Container Domain -From Containers -MasterName 'Domain'

Get-Location MyCity {
	WebServer PrimaryServer 5 5 
}
