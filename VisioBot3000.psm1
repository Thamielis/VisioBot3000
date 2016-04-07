Set-StrictMode -Version Latest
#Need System.Drawing for Colors.
[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | out-null

#module level variables
$Visio=0
$Shapes=@{}
$Stencils=@{}
$updateMode=$false 
$LastDroppedObject=0
$RelativeOrientation='Horizontal'

Function New-VisioApplication{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([switch]$Hide)
    if($PSCmdlet.ShouldProcess('Creating a new instance of Visio','')){
        if ($Hide){
            $script:Visio=New-Object -ComObject Visio.InvisibleApp 
        } else {
            $script:Visio = New-Object -ComObject Visio.Application
        }
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
        $script:updatemode=$True 
    }
}

Function New-VisioDocument{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([string]$Path,
        [string]$From='',
        $Visio=$script:visio,
        [switch]$Update,
    [switch]$landscape,[switch]$portrait)
    if($PSCmdlet.ShouldProcess('Creating a new Visio Document','')){
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
        if($landscape){
            $Visio.ActiveDocument.DiagramServicesEnabled=8
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=2
        } else {
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=1
        }
        $Visio.ActiveDocument.SaveAs($path)
    }
}
Function Get-VisioDocument{
    [CmdletBinding()]
    Param($Visio=$script:Visio)
    return $Visio.ActiveDocument
}

Function New-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([string]$name,
    $Visio=$script:Visio)

    if($PSCmdlet.ShouldProcess('Creating a new Visio Page')){
        $page=$Visio.ActiveDocument.Pages.Add( )
        if($name){
            $page.NameU=$name 
        }
        $page
    }
}
Function Set-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([string]$name,
    $Visio=$script:Visio)
    if($PSCmdlet.ShouldProcess('Switching to a different Visio Page','')){
        $page=get-VisioPage $name
        $Visio.ActiveWindow.Page=$page 
    }
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
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($name)
    if($PSCmdlet.ShouldProcess('Removing page named <$name> or current page','')){
        if ($name) {
            $Visio.ActiveDocument.Pages($name).Delete(0)
        } else {
            $Visio.ActivePage.Delete(0)
        }
    }

}
Function Set-VisioPageLayout{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([switch]$landscape,[switch]$portrait)
    if($PSCmdlet.ShouldProcess('Visio','Switch page layout')){
        if($landscape){
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=2
        } else {
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=1
        }
    }
}

Function New-VisioShape{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($master,$label,$x,$y,$name)
    if($PSCmdlet.ShouldProcess('Visio','Drop shape on the page')){
        if($master -is [string]){
            $master=$script:Shapes[$master]
        }
        if(!$name){
            $name=$label
        }
 
        $p=get-VisioPage
        if($updateMode){
            $DroppedShape=$p.Shapes | Where-Object {$_.Name -eq $label}
        }
        if(-not (get-variable DroppedShape -Scope Local -ErrorAction Ignore) -or $DroppedShape -eq $null){
            if(-not $x){
                $RelativePosition=Get-NextShapePosition
                $x=$RelativePosition.X
                $y=$RelativePosition.Y
            }
            $DroppedShape=$p.Drop($master.PSObject.BaseObject,$x,$y)
            $Script:LastDroppedObject=$DroppedShape
            $DroppedShape.Name=$name
        } else {
            write-verbose "Existing shape <$label> found"
        }
        $DroppedShape.Text=$label
        New-Variable -Name $name -Value $DroppedShape -Scope Global -Force
        write-output $DroppedShape
    }
}

Function New-VisioRectangle{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($x0,$y0,$x1,$y1)
    if($PSCmdlet.ShouldProcess('Visio','Draw a rectangle on the page')){    
        $p=get-visioPage
        $p.DrawRectangle($x0,$y0,$x1,$y1)
    }
}

Function New-VisioConnector{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($from,
        $to,
        $name,
        [System.Drawing.Color]$color,
        [switch]$Arrow,
        [switch]$bidirectional,
    $label)
    if($PSCmdlet.ShouldProcess('Visio','Connect shapes with a connector')){
        $CurrentPage=Get-VisioPage
        if($from -is [string]){
            $from=$CurrentPage.Shapes[$from]
        }
        if($to -is [string]){
            $to=$CurrentPage.Shapes[$to]
        }
        if(!$name){
            $Name='{0}_{1}_{2}' -f $label,$from.Name,$to.Name
        } 
        if($updatemode){
            $connector=$CurrentPage.Shapes | Where-Object {$_.Name -eq $name}
        }
        if (-not (get-variable Connector -Scope Local -ErrorAction Ignore)){
            $from.AutoConnect($to,0)
            $connector=$CurrentPage.Shapes('Dynamic Connector')| Select-Object -first 1
            $connector.Name=$name
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
}

Function New-VisioContainer{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param( [string]$name,
        [Scriptblock]$contents,
        $shape,
    $label)
    if(!$label){
        $label=$name
    }
     if($PSCmdlet.ShouldProcess('Visio','Drop a container around shapes')){
        $page=Get-VisioPage
        if($contents){
            [array]$containedObjects=& $contents
            $firstShape=$containedObjects[0]
            if($updatemode){
                $droppedContainer=$page.Shapes | Where-Object {$_.Name -eq $label}
            }
        
            if(get-variable droppedContainer -Scope Local -ErrorAction Ignore){
                If($droppedContainer.ContainerProperties.GetMemberShapes(16+2) -notcontains $firstShape.ID){
                    $droppedcontainer.ContainerProperties.AddMember($firstShape,2)
                }
            } else {
                $sel=New-VisioSelection $firstShape -Visible
                $droppedContainer=$page.DropContainer($shape,$page.Application.ActiveWindow.Selection)
                $Script:LastDroppedObject=$droppedContainer
                $droppedContainer.Name=$name
            } 
            $droppedContainer.ContainerProperties.SetMargin($vis.PageUnits, 0.25)
            $containedObjects | select-object -Skip 1 | % { 
                if(-not $updatemode -or ($droppedContainer.ContainerProperties.GetMemberShapes(16+2) -notcontains $_.ID)){
                    $droppedcontainer.ContainerProperties.AddMember($_,1)
                }
            }        
            $droppedContainer.ContainerProperties.FitToContents()
            $droppedContainer.Text=$label
            $droppedContainer

        }
        New-Variable -Name $name -Value $droppedContainer -Scope Global -Force
    }
}

Function Register-VisioBuiltinStencil{
    [CmdletBinding()]
    Param([ValidateSet('Backgrounds','Borders','Containers','Callouts','Legends')]
        [string]$BuiltInStencil,
    [String]$Name)
    $stencilID=@('Backgrounds','Borders','Containers','Callouts','Legends').IndexOf($BuiltInStencil)
    $stencilPath=$Visio.GetBuiltInStencilFile($stencilID,$vis.MSDefault)
    Register-VisioStencil -Path $stencilPath -Name $Name 
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
 

    $newShape=$stencils[$StencilName].Masters | Where-Object {$_.Name -eq $masterName}
    $script:Shapes[$name]=$newshape
    $outerName=$name 
    new-item -Path Function:\ -Name "global`:$outername" -value {param($label, $x,$y, $name) $shape=get-visioshape $outername; New-VisioShape $shape $label $x $y -name $name}.GetNewClosure() -force  | out-null

}
Function Register-VisioContainer{
    [CmdletBinding()]
    Param([string]$name,
        [Alias('From')][string]$StencilName,
    [string]$masterName)
 

    $newShape=$stencils[$StencilName].Masters | Where-Object {$_.Name -eq $masterName}
    $script:Shapes[$name]=$newshape
    $outerName=$name
    new-item -Path Function:\ -Name "global`:$outername" -value {param($label,$contents,$name) $shape=get-visioshape $outername; New-VisioContainer  $label $contents $shape $name}.GetNewClosure() -force  | out-null

}
Function Register-VisioConnector{
    [CmdletBinding()]
    Param([string]$name,
        [System.Drawing.Color]$color,
        [switch]$Arrow,
    [switch]$bidirectional)
    new-item -Path Function:\ -Name "global`:$name" -value {param($from,$to,$label) New-VisioConnector $from $to $name $color -Arrow:$Arrow.IsPresent -bidirectional:$bidirectional.IsPresent $label}.GetNewClosure() -force  | out-null
}


Function Get-VisioShape{
    [CmdletBinding()]
    Param([string]$name)
    $script:Shapes[$name]
}
Function New-VisioHyperlink{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($shape,
    $link)
    if($PSCmdlet.ShouldProcess('Visio','Create a hyperlink on a shape')){
        $CurrentPage=Get-VisioPage
        if($shape -is [string]){
            $shape=$CurrentPage.Shapes[$shape]
        }
        $linkObject=$shape.AddHyperLink()
        $linkObject.Address=$link
    }
}

Function New-VisioSelection{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($Objects,[switch]$Visible)
    if($PSCmdlet.ShouldProcess('Visio','Create a selection object')){
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
}


Function Set-VisioShapeData{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param($Shape,
        $Name,
    $Value)
        if($PSCmdlet.ShouldProcess('Visio','Set a value for a custom shape data element')){
        $shape.Cells("Prop.$Name").Formula="=`"$value`""
    }
}

Function Get-VisioShapeData{
    [CmdletBinding()]
    Param($Shape,
    $Name)
    if($PSCmdlet.ShouldProcess('Visio','Retrieve the value from a custom shape data element')){
        $shape.Cells("Prop.$Name").Formula.TrimStart('"').TrimEnd('"') 
    }
}

Function Complete-Diagram{
    [CmdletBinding()]
    Param([switch]$Close)
    $script:updateMode=$false
    $Visio.ActiveDocument.Save() 
    if($Close){
        $Visio.Quit()
    }
}

Function New-VisioLayer{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([string]$LayerName,$Contents,[switch]$Preserve)
    if($PSCmdlet.ShouldProcess('Visio','Create a layer on the current page')){
        if($Preserve){
            $AddOption=1
        } else {
            $AddOption=0
        }
        $p=$Visio.ActivePage
        $layer=$p.Layers | Where-Object {$_.Name -eq $LayerName}
        if ($layer -eq $null){
            $layer=$p.Layers.Add($LayerName) 
        }
        if ($contents -is [scriptblock]){
            $Contents = & $contents
        }
        foreach($item in [array]$Contents){
            if($item -is [string]){
                $item=$p.Shapes[$item] 
            }
            $layer.Add($item,$AddOption)
        }
    }
}

Function Get-NextShapePosition{
    [CmdletBinding()]
    Param()
    if($LastDroppedObject -eq 0){
        #nothing dropped yet, start at top-left-ish
        return @{X=1;Y=10}
    } else {
        if($RelativeOrientation -eq 'Horizontal'){
            $x=$LastDroppedObject.Cells('PinX').ResultIU + $LastDroppedObject.Cells('Width').ResultIU + 0.25
            $y=$LastDroppedObject.Cells('PinY').ResultIU 
        } else {
            $x=$LastDroppedObject.Cells('PinX').ResultIU 
            $y=$LastDroppedObject.Cells('PinY').ResultIU - $LastDroppedObject.Cells('Height').ResultIU - 0.25
        }
        Return @{X=$x;Y=$y}
    }
}

Function Set-RelativePositionDirection{
    [CmdletBinding(SupportsShouldProcess=$true)]
    Param([ValidateSet('Horizontal','Vertical')]$Orientation)
    if($PSCmdlet.ShouldProcess('Visio','Sets VisioBot''s orientation for relative positioning')){
        $script:RelativeOrientation=$Orientation
    }
}

#Aliases
New-Alias -Name Diagram -Value New-VisioDocument
New-Alias -Name Stencil -Value Register-VisioStencil
New-Alias -Name Shape -Value Register-VisioShape
New-Alias -Name Container -Value Register-VisioContainer
New-Alias -Name Connector -Value Register-VisioConnector
New-Alias -Name HyperLink -Value New-VisioHyperlink
New-Alias -Name Layer -value New-VisioLayer