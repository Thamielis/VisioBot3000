$Visio=0
$Shapes=@{}
$Stencils
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
    Param([Alias('From')][string]$path,
    $Visio=$script:visio)
    Open-VisioDocument $path
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
    Param($master,$x0,$y0 )
    if($master -is [string]){
        $master=$script:Shapes[$master]
    }
    $p=get-visioPage
    $p.Drop($master.PSObject.BaseObject,$x0,$y0)
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
Function Import-VisioBuiltinStencil{
    if(!$script:BuiltInStencil){
        $window=(Get-VisioApplication).ActiveWindow
        $script:BuiltInStencil = $Visio.Documents.OpenEx($Visio.GetBuiltInStencilFile(2,1),64)
        $window.Activate()
    }
    $script:BuiltInStencil
}
Function Import-VisioStencil{
    Param([string]$path,
          [string]$Name)
        $stencil=$Visio.Documents.OpenEx($path,$vis.OpenHidden)
        $script:stencils[$Name]=$stencil  
}

Function Register-VisioShape{
    Param([string]$masterName,
        [string]$path,
        [string]$nickname,
        [switch]$builtin,
        [HashTable]$Map,
    $Visio=$script:Visio)
    if(!$nickname){
        $nickname=$masterName
    }
    if(!$map){
        $map=@{$nickName=$masterName}
    }

    if($builtin){
        $stencil=Import-VisioBuiltinStencil
    } else {
        $stencil=Import-VisioStencil $path 
    }
    foreach($nickname in $map.Keys){
        $newShape=$stencil.Masters | Where-Object Name -eq $map.$nickname
        $script:Shapes[$nickname]=$newshape
        new-item -Path Function:\ -Name "global`:$nickname" -value {param($x,$y) $shape=get-visioshape $nickname; $p=get-visiopage;$p.Drop($shape.PSObject.BaseObject,$x,$y)}.GetNewClosure() -force  
    } 

}



Function Get-VisioShape{
    Param([string]$name)
    $script:SavedShapes[$name]
}

#Aliases
New-Alias -Name Diagram -Value New-VisioDocument


Export-ModuleMember *-* -alias *