<#
        .SYNOPSIS 
        Reads a psd1 file with settings for the current diagram.  

        .DESCRIPTION
        Reads a psd1 file with settings for the current diagram.  Should be called after a diagram has been opened or created.  
        Currently supported sections are:
        StenciPaths - list of paths to be added to the stencilpath
        Stencils - hashtable with name=nickname of stencil, value=filename to stencil
        Shapes - hashtable with name=nickname, value=array with stencilname and mastername
        Containers -  hashtable with name=nickname, value=array with stencilname and mastername
        Connectors - hashtable with name=nickname, value=hashtable of parameters splatted to register-visioconnector

        .PARAMETER Path
        The path to the psd1 file.  Must be a full path, not just a filename.


        .INPUTS
        You cannot pipe anything to Import-VisioSettings

        .OUTPUTS
        None

        .EXAMPLE
        Import-VisioSettings c:\Config\DepartmentalDiagramSettings.psd1


#>
function Import-VisioSettings{
[CmdletBinding()]
Param([string]$path)
    $dir=Split-Path -Path $path -Parent 
    $file=split-path -Path $path -leaf
    $settings=Import-LocalizedData -FileName $file -BaseDirectory $dir  
    if($settings.StencilPaths){
        $settings.StencilPaths  | foreach-object {Add-StencilSearchPath -Path $_}
    }

    if($settings.Stencils){
        $Settings.Stencils.GetEnumerator() | foreach{Register-VisioStencil -Name $_.Key -Path $_.Value}
    }

    if($settings.Shapes){
        $Settings.Shapes.GetEnumerator() | foreach{Register-VisioShape -Name $_.Key -From $_.Value[0] -MasterName $_.Value[1]}
    }
   if($settings.Containers){
        $Settings.Containers.GetEnumerator() | foreach{Register-VisioContainer -Name $_.Key -From $_.Value[0] -MasterName $_.Value[1]}
    }
  if($settings.Connectors){
        
        $Settings.Connectors.GetEnumerator() | foreach{$options=$_.Value;Register-VisioConnector -Name $_.Key @options}
    }


}