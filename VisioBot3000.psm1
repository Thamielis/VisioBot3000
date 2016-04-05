Set-StrictMode -Version Latest

#module level variables
$Visio=0
$Shapes=@{}
$Stencils=@{}


Function New-VisioApplication{
[CmdletBinding()]
    Param([switch]$Hide)
    if ($Hide){
        $script:Visio=New-Object -ComObject Visio.InvisibleApp 
    } else {
        $script:Visio = New-Object -ComObject Visio.Application
    }
}
Function Get-VisioApplication{
[CmdletBinding()]
Param()
    if(!$script:Visio){
        New-VisioApplication
    }
    return $Visio
} 

Function Open-VisioDocument{
[CmdletBinding()]
Param([string]$path,
    $Visio=$script:Visio,
    [switch]$Update)
    if(!$Visio){
        New-VisioApplication
        $Visio=$script:Visio
    }
    $documents = $Visio.Documents
    $document = $documents.Add($path)
    if($Update){
        $PSDefaultParameterValues['*-Visio*:Default']=$true
    }
}

Function New-VisioDocument{
[CmdletBinding()]
    Param([string]$Path,
    [string]$From='',
    $Visio=$script:visio,
    [switch]$Update)
    if(!$Visio){
        New-VisioApplication
        $Visio=$script:Visio
    }
    if($Update){
        if($From -ne ''){
            Write-Warning 'New-VisioDocument: -From ignored when -Update is present'
        }
        Open-VisioDocument $path -Update
    } else {
        Open-VisioDocument $From
    }
    $Visio.ActiveDocument.SaveAs($path)
}
Function Get-VisioDocument{
[CmdletBinding()]
    Param($Visio=$script:Visio)
    return $Visio.ActiveDocument
}

Function New-VisioPage{
[CmdletBinding()]
    Param([string]$name,
    $Visio=$script:Visio)

    $page=$Visio.ActiveDocument.Pages.Add( )
    if($name){
        $page.NameU=$name 
    }
    $page
}
Function Set-VisioPage{
[CmdletBinding()]
    Param([string]$name,
    $Visio=$script:Visio)
    $page=get-VisioPage $name
    $Visio.ActiveWindow.Page=$page 
}

Function Get-VisioPage{
[CmdletBinding()]
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
[CmdletBinding()]
    Param($name)
    if ($name) {
        $Visio.ActiveDocument.Pages($name).Delete(0)
    } else {
        $Visio.ActivePage.Delete(0)
    }

}
Function Set-VisioPageLayout{
[CmdletBinding()]
    Param([switch]$landscape,[switch]$portrait)
    if($landscape){
        $Visio.ActiveDocument.PrintLandscape=(1)
    } else {
        $Visio.ActiveDocument.PrintLandscape=(0)
    }
}

Function New-VisioShape{
[CmdletBinding()]
    Param($master,$label,$x=0,$y=0,[switch]$Update )
    if($master -is [string]){
        $master=$script:Shapes[$master]
    }
    $p=get-VisioPage
    if($update){
      $DroppedShape=$p.Shapes | Where-Object {$_.Name -eq $label}
    }
    if(-not (get-variable DroppedShape -Scope Local -ErrorAction Ignore)){
        $DroppedShape=$p.Drop($master.PSObject.BaseObject,$x,$y)
    } else {
        write-verbose "Existing shape <$label> found"
    }
    $DroppedShape.Name=$label
    $DroppedShape.Text=$label
    New-Variable -Name $label -Value $DroppedShape -Scope Global -Force
    write-output $DroppedShape
}

Function New-VisioRectangle{
[CmdletBinding()]
    Param($x0,$y0,$x1,$y1)
    $p=get-visioPage
    $p.DrawRectangle($x0,$y0,$x1,$y1)
}

Function New-VisioConnector{
[CmdletBinding()]
    Param($from,
          $to,
          $Label,
          [System.Drawing.Color]$color,
          [switch]$Arrow,
          [switch]$bidirectional,
          [switch]$Update)
    $CurrentPage=Get-VisioPage
    if($from -is [string]){
        $from=$CurrentPage.Shapes[$from]
    }
    if($to -is [string]){
        $to=$CurrentPage.Shapes[$to]
    }
    $Name='{0}_{1}_{2}' -f $label,$from.Name,$to.Name
    if($update){
      $connector=$p.Shapes | Where-Object {$_.Name -eq $name}
    }
    if (-not (get-variable Connector -Scope Local -ErrorAction Ignore)){
        $from.AutoConnect($to,0)
        $connector=$CurrentPage.Shapes('Dynamic Connector')| Select-Object -first 1
        $connector.Name='{0}_{1}_{2}' -f $label,$from.Name,$to.Name
   }
    $connector.Text=$label
    $connector.CellsU('LineColor').Formula="rgb($($color.R),$($color.G),$($color.B))"
    $connector.CellsSRC(1,23,10) = 16
    $connector.CellsSRC(1,23,19) = 1 

    if($Arrow){
         $connector.Cells('EndArrow')=5
         if($bidirectional){ 
            $connector.Cells(‘BeginArrow')=5
         } else {
            
         }
    } else {
        $connector.Cells('EndArrow')=0
        $connector.Cells('BeginArrow')=0
    }
}

Function New-VisioContainer{
[CmdletBinding()]
    Param( [string]$label,
        [Scriptblock]$contents,
    $shape,
    [switch]$update)
    $page=Get-VisioPage
    if($contents){
        [array]$containedObjects=& $contents
        $firstShape=$containedObjects[0]
        if($update){
           $droppedContainer=$page.Shapes | Where-Object {$_.Name -eq $label}
        }
        if(get-variable droppedContainer -Scope Local -ErrorAction Ignore){
          If($droppedContainer.ContainerProperties.GetMemberShapes(16+2) -notcontains $firstShape.ID){
            $droppedcontainer.ContainerProperties.AddMember($firstShape,2)
          }
        } else {
            $droppedContainer=$page.DropContainer($shape,$firstShape)
        } 
        $droppedContainer.ContainerProperties.SetMargin($vis.PageUnits, 0.25)
        $containedObjects | select-object -Skip 1 | % { 
            if(-not $update -or ($droppedContainer.ContainerProperties.GetMemberShapes(16+2) -notcontains $_.ID)){
              $droppedcontainer.ContainerProperties.AddMember($_,1)
            }
        }        
        $droppedContainer.ContainerProperties.FitToContents()
        $droppedContainer.Text=$label
        $droppedContainer.Name=$label
        $droppedContainer

    }
    New-Variable -Name $label -Value $droppedContainer -Scope Global -Force

}

Function Register-VisioBuiltinStencil{
[CmdletBinding()]
Param([ValidateSet('Backgrounds','Borders','Containers','Callouts','Legends')]
[string]$BuiltInStencil,
[String]$Name)
    $stencilID=@('Backgrounds','Borders','Containers','Callouts','Legends').IndexOf($BuiltInStencil)
    $stencilPath=$Visio.GetBuiltInStencilFile($stencilID,$vis.MSDefault)
    Import-VisioStencil -Path $stencilPath -Name $Name 
}
Function Register-VisioStencil{
[CmdletBinding()]
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
[CmdletBinding()]
    Param([string]$name,
         [Alias('From')][string]$StencilName,
         [string]$masterName)
 

        $newShape=$stencils[$StencilName].Masters | Where-Object Name -eq $masterName
        $script:Shapes[$name]=$newshape
        new-item -Path Function:\ -Name "global`:$name" -value {param($label, $x,$y) $shape=get-visioshape $name; New-VisioShape $shape $label $x $y}.GetNewClosure() -force  | out-null

}
Function Register-VisioContainer{
[CmdletBinding()]
    Param([string]$name,
         [Alias('From')][string]$StencilName,
         [string]$masterName)
 

        $newShape=$stencils[$StencilName].Masters | Where-Object Name -eq $masterName
        $script:Shapes[$name]=$newshape
        new-item -Path Function:\ -Name "global`:$name" -value {param($label,$contents) $shape=get-visioshape $name; New-VisioContainer  $label $contents $shape}.GetNewClosure() -force  | out-null

}
Function Register-VisioConnector{
[CmdletBinding()]
    Param([string]$name,
          [System.Drawing.Color]$color,
          [switch]$Arrow,
          [switch]$bidirectional)
         new-item -Path Function:\ -Name "global`:$name" -value {param($from,$to) New-VisioConnector $from $to $name $color -Arrow:$Arrow.IsPresent -bidirectional:$bidirectional.IsPresent}.GetNewClosure() -force  | out-null
}


Function Get-VisioShape{
[CmdletBinding()]
    Param([string]$name)
    $script:Shapes[$name]
}
Function New-VisioHyperlink{
[CmdletBinding()]
Param($shape,
      $link)
    $CurrentPage=Get-VisioPage
    if($shape -is [string]){
        $shape=$CurrentPage.Shapes[$shape]
    }
    $linkObject=$shape.AddHyperLink()
    $linkObject.Address=$link

}

Function New-VisioSelection{
[CmdletBinding()]
    Param($Objects,[switch]$Visible)
    $V=Get-VisioApplication
    $sel=$v.ActiveWindow.Selection
    if($visible){
        $sel=$v.ActiveWindow
    }
    $sel.DeselectAll()
    $CurrentPage=Get-VisioPage
    foreach($o in $objects){
        if($o -is [string]){
         $o=$CurrentPage.Shapes[$o]
       }
       $sel.Select($o,2)
    }
    $sel
}


Function Set-VisioShapeData{
[CmdletBinding()]
    Param($Shape,
          $Name,
          $Value)
    $shape.Cells("Prop.$Name").Formula="`"$value`""
}

Function Get-VisioShapeData{
[CmdletBinding()]
    Param($Shape,
          $Name,
          $Value)
    $shape.Cells("Prop.$Name").Formula="`"$value`""
}

Function Complete-Diagram{
[CmdletBinding()]
    Param([switch]$Close)

    $Visio.ActiveDocument.Save() 
    if($Close){
        $Visio.Quit()
    }
}

#Aliases
New-Alias -Name Diagram -Value New-VisioDocument
New-Alias -Name Stencil -Value Register-VisioStencil
New-Alias -Name Shape -Value Register-VisioShape
New-Alias -Name Container -Value Register-VisioContainer
New-Alias -Name Connector -Value Register-VisioConnector
New-Alias -Name HyperLink -Value New-VisioHyperlink