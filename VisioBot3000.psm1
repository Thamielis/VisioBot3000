Set-StrictMode -Version Latest
$Visio=0
$Shapes=@{}
$Stencils=@{}
Function New-VisioApplication{
    Param([switch]$Hide)
    if ($Hide){
        $script:Visio=New-Object -ComObject Visio.InvisibleApp 
    } else {
        $script:Visio = New-Object -ComObject Visio.Application
    }
}
Function Get-VisioApplication{
    if(!$script:Visio){
        New-VisioApplication
    }
    return $Visio
} 

Function Open-VisioDocument{
    Param([string]$path,
    $Visio=$script:Visio)
    $documents = $Visio.Documents
    $document = $documents.Add($path)

}

Function New-VisioDocument{
    Param([string]$Path,
    [string]$From='',
    $Visio=$script:visio)
    Open-VisioDocument $From
    $Visio.ActiveDocument.SaveAs($path)
}
Function Get-VisioDocument{
    Param($Visio=$script:Visio)
    return $Visio.ActiveDocument
}

Function New-VisioPage{
    Param([string]$name,
    $Visio=$script:Visio)

    $page=$Visio.ActiveDocument.Pages.Add( )
    if($name){
        $page.NameU=$name 
    }
    $page
}
Function Set-VisioPage{
    Param([string]$name,
    $Visio=$script:Visio)
    $page=get-VisioPage $name
    $Visio.ActiveWindow.Page=$page 
}

Function Get-VisioPage{
    Param($name)
    if ($name) {
        try {
            $Visio.ActiveDocument.Pages($name) 
        } catch {
            write-warning "$name not found"
        }
    } else {
        $Visio.ActivePage
    }
}

Function Remove-VisioPage{
    Param($name)
    if ($name) {
        $Visio.ActiveDocument.Pages($name).Delete(0)
    } else {
        $Visio.ActivePage.Delete(0)
    }

}
Function Set-VisioPageLayout{
    Param([switch]$landscape,[switch]$portrait)
    if($landscape){
        $Visio.ActiveDocument.PrintLandscape=(1)
    } else {
        $Visio.ActiveDocument.PrintLandscape=(0)
    }
}

Function New-VisioShape{
    Param($master,$x,$y )
    if($master -is [string]){
        $master=$script:Shapes[$master]
    }
    $p=get-visioPage
    $p.Drop($master.PSObject.BaseObject,$x,$y)
}

Function New-VisioRectangle{
    Param($x0,$y0,$x1,$y1)
    $p=get-visioPage
    $p.DrawRectangle($x0,$y0,$x1,$y1)
}

Function New-VisioConnection{
    Param($from,$to)
    $from.AutoConnect($to,0)
    (get-visiopage).Shapes('Dynamic Connector')| Select-Object -first 1

}

Function New-VisioContainer{
    Param( [string]$label,
        [Scriptblock]$contents,
    $shape)
    $page=Get-VisioPage
    if(!$shape){
        Load-VisioBuiltinStencil
        $shape = $script:BuiltInStencil.Masters.ItemFromID(2)
    }
    if($contents){
        $containedObjects=& $contents

        $droppedContainer=$page.Drop($shape,0,0)
        $containedObjects | % { 
            $droppedcontainer.ContainerProperties.AddMember($_,1)
        }
        $droppedContainer.ContainerProperties.FitToContents()
        $droppedContainer.Text=$label
        $droppedContainer.Name=$label
        $droppedContainer
  
    }
}
Function Register-VisioBuiltinStencil{
Param([ValidateSet('Backgrounds','Borders','Containers','Callouts','Legends')]
[string]$BuiltInStencil,
[String]$Name)
    $stencilID=@('Backgrounds','Borders','Containers','Callouts','Legends').IndexOf($BuiltInStencil)
    $stencilPath=$Visio.GetBuiltInStencilFile($stencilID,$vis.MSDefault)
    Import-VisioStencil -Path $stencilPath -Name $Name 
}
Function Register-VisioStencil{
    Param([string]$Name,
          [Alias('From')][string]$Path,
          [switch]$BuiltIn)
        if($BuiltIn){
            if(!$Path){
                $Path=$Name
            }
            Register-VisioBuiltinStencil -BuiltInStencil $Path -Name $Name 
        } else {
            $stencil=$Visio.Documents.OpenEx($Path,$vis.OpenHidden)
            $script:stencils[$Name]=$stencil
        }  
}

Function Register-VisioShape{
    Param([string]$name,
         [Alias('From')][string]$StencilName,
         [string]$masterName)
 

        $newShape=$stencils[$StencilName].Masters | Where-Object Name -eq $masterName
        $script:Shapes[$name]=$newshape
        new-item -Path Function:\ -Name "global`:$name" -value {param($x,$y) $shape=get-visioshape $name; $p=get-visiopage;$p.Drop($shape.PSObject.BaseObject,$x,$y)}.GetNewClosure() -force  | out-null

}



Function Get-VisioShape{
    Param([string]$name)
    $script:SavedShapes[$name]
}

#Aliases
New-Alias -Name Diagram -Value New-VisioDocument
New-Alias -Name Stencil -Value Register-VisioStencil
New-Alias -Name Shape -Value Register-VisioShape
