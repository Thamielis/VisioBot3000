$application=0
$savedShapes=@{}
function New-VisioApplication{
Param([switch]$hide)
$script:application = New-Object -ComObject Visio.Application
}
function Get-VisioApplication{
  return $application
} 

function Open-VisioDocument{
Param($path)
$documents = $application.Documents
$document = $documents.Add($path)

}

function New-VisioDocument{
Open-VisioDocument ''
}
function Get-VisioDocument{
return $application.ActiveDocument
}

function New-VisioPage{
Param($name)
   $page=$application.ActiveDocument.Pages.Add( )
   if($name){
      $page.NameU=$name 
   }
   $page
}
function Set-VisioPage{
Param($name)
  $page=get-VisioPage $name
  $application.ActiveWindow.Page=$page 
}

function Get-VisioPage{
Param($name)
    if ($name) {
       try {
       $application.ActiveDocument.Pages($name) 
       } catch {
         write-warning "$name not found"
       }
    } else {
       $application.ActivePage
    }
}

function Remove-VisioPage{
Param($name)
   if ($name) {
       $application.ActiveDocument.Pages($name).Delete(0)
    } else {
       $application.ActivePage.Delete(0)
    }

}
function Set-VisioPageLayout{
Param([switch]$landscape,[switch]$portrait)
    if($landscape){
     $application.ActiveDocument.PrintLandscape=(1)
    } else {
     $application.ActiveDocument.PrintLandscape=(0)
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
    (get-visiopage).Shapes('Dynamic Connector')| select -first 1

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
        $script:BuiltInStencil = $Application.Documents.OpenEx($Application.GetBuiltInStencilFile(2,1),64)
        $window.Activate()
    }
    $script:BuiltInStencil
}
function Import-VisioStencil{
Param($path)
    if(!$script:BuiltInStencil){
        $window=(Get-VisioApplication).ActiveWindow
        $Application.Documents.OpenEx($path,64)
        $window.Activate()
    }
    $$
}

function Register-VisioShape{
Param($masterName,$path,$nickname,[switch]$builtin,$Map)
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
  $newShape=$stencil.Masters | where Name -eq $map.$nickname
  $script:SavedShapes[$nickname]=$newshape
  new-item -Path Function:\ -Name "global`:$nickname" -value {param($x,$y) $shape=get-visioshape $nickname; $p=get-visiopage;$p.Drop($shape.PSObject.BaseObject,$x,$y)}.GetNewClosure() -force  

  } 

}



function Get-VisioShape{
Param($name)
  $script:SavedShapes[$name]
}

export-modulemember *-*