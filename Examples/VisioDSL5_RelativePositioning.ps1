﻿#clean-up because I'm re-running this over and over
stop-process -Name VISIO -ea SilentlyContinue
remove-item c:\temp\testvisio5.vsdx -ea SilentlyContinue
import-module VisioBot3000 -Force

Diagram C:\temp\TestVisio5.vsdx -From C:\temp\IntegrationDiagram.vstx

# Define shapes, containers, and connectors for the diagram
Stencil Servers -From C:\temp\SERVER_U.vssx
Stencil Containers -From C:\temp\MyContainers.vssx 
Shape WebServer -From Servers -MasterName 'Web Server'
Container Location -From Containers -MasterName 'Location'

Location Datacenter {
       WebServer PrimaryServer
       WebServer SecondaryServer
       WebServer ThirdServer
}


Legend @{
            'Information/CreatedBy/Name'='Mike Shepard - the boss!';
            'Information/LastUpdateBy/Name'='Mike Shepard - the boss2!';
            'Title/Title'='VisioBot3000 DSL Example';
            'Title/SubTitle'='Relative positioning and legend'}


Complete-Diagram 