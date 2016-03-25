$Visio=0
$savedShapes=@{}
function New-VisioApplication{
Param([switch]$hide)
$script:Visio = New-Object -ComObject Visio.Application
}
function Get-VisioApplication{
  if(!$module:Visio){
    New-VisioApplication
  }
  return $Visio
} 

function Open-VisioDocument{
Param([string]$path,
      $Visio=$module:Visio)
$documents = $Visio.Documents
$document = $documents.Add($path)

}

function New-VisioDocument{
Param([Alias('From')][string]$path,
      $Visio=$module:visio)
Open-VisioDocument $path
}
function Get-VisioDocument{
Param($Visio=$module:Visio)
    return $Visio.ActiveDocument
}

function New-VisioPage{
Param([string]$name,
      $Visio=$module:Visio)

   $page=$Visio.ActiveDocument.Pages.Add( )
   if($name){
      $page.NameU=$name 
   }
   $page
}
function Set-VisioPage{
Param([string]$name,
      $Visio=$module:Visio)
  $page=get-VisioPage $name
  $Visio.ActiveWindow.Page=$page 
}

function Get-VisioPage{
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

function Remove-VisioPage{
Param($name)
   if ($name) {
       $Visio.ActiveDocument.Pages($name).Delete(0)
    } else {
       $Visio.ActivePage.Delete(0)
    }

}
function Set-VisioPageLayout{
Param([switch]$landscape,[switch]$portrait)
    if($landscape){
     $Visio.ActiveDocument.PrintLandscape=(1)
    } else {
     $Visio.ActiveDocument.PrintLandscape=(0)
    }
}

function New-VisioShape{
Param($master,$x0,$y0 )
   if($master -is [string]){
     $master=$script:SavedShapes[$master]
   }
   $p=get-visioPage
   $p.Drop($master.PSObject.BaseObject,$x0,$y0)
}

function New-VisioRectangle{
    Param($x0,$y0,$x1,$y1)
    $p=get-visioPage
    $p.DrawRectangle($x0,$y0,$x1,$y1)
}

function New-VisioConnection{
Param($from,$to)
    $from.AutoConnect($to,0)
    (get-visiopage).Shapes('Dynamic Connector')| Select-Object -first 1

}

function New-VisioContainer{
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
function Import-VisioBuiltinStencil{
    if(!$script:BuiltInStencil){
        $window=(Get-VisioApplication).ActiveWindow
        $script:BuiltInStencil = $Visio.Documents.OpenEx($Visio.GetBuiltInStencilFile(2,1),64)
        $window.Activate()
    }
    $script:BuiltInStencil
}
function Import-VisioStencil{
Param($path)
    if(!$script:BuiltInStencil){
        $window=(Get-VisioApplication).ActiveWindow
        $Visio.Documents.OpenEx($path,64)
        $window.Activate()
    }
    $$
}

function Register-VisioShape{
Param([string]$masterName,
      [string]$path,
      [string]$nickname,
      [switch]$builtin,
      [HashTable]$Map,
      $Visio=$module:Visio)
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
      $script:SavedShapes[$nickname]=$newshape
      new-item -Path Function:\ -Name "global`:$nickname" -value {param($x,$y) $shape=get-visioshape $nickname; $p=get-visiopage;$p.Drop($shape.PSObject.BaseObject,$x,$y)}.GetNewClosure() -force  
  } 

}



function Get-VisioShape{
Param([string]$name)
  $module:SavedShapes[$name]
}

Export-ModuleMember *-*