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


<#
        .SYNOPSIS 
        Starts the visio application (possibly hidden) and stores a reference to the application object

        .DESCRIPTION
        Starts the visio application (possibly hidden) and stores a reference to the application object

        .PARAMETER Hide
        Starts Visio without showing the user interface

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        None

        .EXAMPLE
        New-VisioApplication
        --Visio Pops up--

        .EXAMPLE
        New-VisioApplication -Hide
        --Nothing seems to happen

#>
Function New-VisioApplication{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([switch]$Hide,
          [switch]$Passthru)
    if($PSCmdlet.ShouldProcess('Creating a new instance of Visio','')){
        if ($Hide){
            $script:Visio=New-Object -ComObject Visio.InvisibleApp 
        } else {
            $script:Visio = New-Object -ComObject Visio.Application
        }
        if($Passthru){
            $script:Visio
        }
    }


}


<#
        .SYNOPSIS 
        Ouptuts a reference to the Visio application object

        .DESCRIPTION
        Ouptuts a reference to the Visio application object


        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        Visio.Application
        Visio.InvisibleApp

        .EXAMPLE
        $app=Get-VisioApplication
#>
Function Get-VisioApplication{
    [CmdletBinding()]
    Param()
    if(!$script:Visio){
        New-VisioApplication
    }
    return $Visio
} 


<#
        .SYNOPSIS 
        Opens a visio document

        .DESCRIPTION
        Opens an existing Visio document, a blank Visio Document 

        .PARAMETER Path
        The path to an existing document, and empty string (to create a blank document) or the path to a Visio template (vstx, etc.)

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .PARAMETER Update
        Switch indicating that we're updating a diagram, potentially created with VisioBot3000

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        None

        .EXAMPLE
        Open-VisioDocument 
        --Creates a blank document--

        .EXAMPLE
        Open-VisioDocument .\MySampleVisio.vsdx
        --Opens the named document

        .EXAMPLE
        Open-VisioDocument .\MyVisioTemplate.vstx
        --Creates a Visio template for editing (not a new document based on the template)

#>
Function Open-VisioDocument{
    [CmdletBinding()]
    Param([string]$Path,
        $Visio=$script:Visio,
    [switch]$Update)
    if(!$Visio){
        New-VisioApplication
        $Visio=$script:Visio
    }
    $documents = $Visio.Documents
    $documents.Add($Path) | out-null
    if($Update){
        $script:updatemode=$True 
    }
}


<#
        .SYNOPSIS 
        Creates a new document

        .DESCRIPTION
        Creates a new document

        .PARAMETER Path
        The path you want to save the document to 

        .PARAMETER From
        The path to a template file to create the new document from 

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .PARAMETER Update
        Switch indicating that we're updating a diagram, potentially created with VisioBot3000

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        None

        .EXAMPLE
        New-VisioDocument 
        --Creates a blank document--

        .EXAMPLE
        New-VisioDocument .\MySampleVisio.vsdx
        --Opens the named document

        .EXAMPLE
        New-VisioDocument .\MyVisioTemplate.vstx
        --Creates a new document based on a Visio template  

#> 
function New-VisioDocument{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$Path,
        [string]$From='',
        $Visio=$script:visio,
        [switch]$Update,
    [switch]$Landscape,[switch]$portrait)
    if($PSCmdlet.ShouldProcess('Creating a new Visio Document','')){
        if(!$Visio){
            New-VisioApplication
            $Visio=$script:Visio
        }
        if($Update){
            if($From -ne ''){
                Write-Warning 'New-VisioDocument: -From ignored when -Update is present'
            }
            Open-VisioDocument $Path -Update
        } else {
            Open-VisioDocument $From

        }
        
        if($Landscape){
            $Visio.ActiveDocument.DiagramServicesEnabled=8
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=2
        } elseif ($portrait) {
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=1
        }
        $Visio.ActiveDocument.SaveAs($Path)
    }
}

<#
        .SYNOPSIS 
        Outputs the active Visio document

        .DESCRIPTION
        Outputs the active Visio document

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .INPUTS
        None. You cannot pipe objects to Get-VisioDocument.

        .OUTPUTS
        visio.Document

        .EXAMPLE
        $doc=Get-VisioDocument

#>
Function Get-VisioDocument{
    [CmdletBinding()]
    Param($Visio=$script:Visio)
    return $Visio.ActiveDocument
}


Function New-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$Name,
    $Visio=$script:Visio)

    if($PSCmdlet.ShouldProcess('Creating a new Visio Page')){
        $page=$Visio.ActiveDocument.Pages.Add( )
        if($Name){
            $page.NameU=$Name 
        }
        $page
    }
}

<#
        .SYNOPSIS 
        Change the active page in Visio

        .DESCRIPTION
        Changes the active page in Visio to the page named in the parameter

        .PARAMETER Name
        Page name in the Visio document which you want to switch to

        .PARAMETER Visio
        Optional reference to a Visio Application (used if writing to multiple diagrams at the same time?)

        .INPUTS
        None. You cannot pipe objects to Set-VisioPage

        .OUTPUTS
        None

        .EXAMPLE
        Set-VisioPage -Page 'Page-3'


#>
function Set-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$Name,
    $Visio=$script:Visio)
    if($PSCmdlet.ShouldProcess('Switching to a different Visio Page','')){
        $page=get-VisioPage $Name
        $Visio.ActiveWindow.Page=$page 
    }
} 

<#
        .SYNOPSIS 
        Returns a visio page

        .DESCRIPTION
        Returns either the named page or the active page if nothing was named.

        .PARAMETER Name
        The name of the page you want.  If you don't supply a name, the active page will be output.

        .INPUTS
        None. You cannot pipe objects to Get-VisioPage.

        .OUTPUTS
        Visio.Page

        .EXAMPLE
        $activePage=get-VisioPage
        #Returns the active page

        .EXAMPLE
        get-VisioPage 'Page-3'
        #returns the page named 'Page-3'


#>
Function Get-VisioPage{
    [CmdletBinding()]
    Param($Name)
    if ($Name) {
        try {
            $Visio.ActiveDocument.Pages($Name) 
        } catch {
            write-warning "$Name not found"
        }
    } else {
        $Visio.ActivePage
    }
}


<#
        .SYNOPSIS 
        Deletes a page from Visio

        .DESCRIPTION
        Deletes a named page or the active page if no page is named.

        .PARAMETER Name
        The name of the page to remove.  If no page is named, the active page is removed.

        .PARAMETER Parameter2
        Describe Parameter1

        .INPUTS
        What can be piped in
        None. You cannot pipe objects to Remove-VisioPage

        .OUTPUTS
        None

        .EXAMPLE
        Remove-VisioPage 'Page-3'
        #removes page 3

        .EXAMPLE
        Remove-VisioPage
        #removes the active page

#>
Function Remove-VisioPage{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Name)
    if($PSCmdlet.ShouldProcess('Removing page named <$Name> or current page','')){
        if ($Name) {
            $Visio.ActiveDocument.Pages($Name).Delete(0)
        } else {
            $Visio.ActivePage.Delete(0)
        }
    }

}


<#
        .SYNOPSIS 
        Switches the page orientation

        .DESCRIPTION
        Set the page orientation to either Landscape or Portrait

        .PARAMETER Landscape
        Changes the page orientation to Landscape

        .PARAMETER Portrait
        Changes the page orientation to Portrait

        .INPUTS
        None. You cannot pipe objects to Set-VisioPageLayout

        .OUTPUTS
        None

        .EXAMPLE
        Set-VisioPageLayout -Portrait

        .EXAMPLE
        Set-VisioPageLayout -Landscape

#>
Function Set-VisioPageLayout{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([switch]$Landscape,[switch]$Portrait)
    if($PSCmdlet.ShouldProcess('Visio','Switch page layout')){
        if($Landscape){
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=2
        } else {
            $Visio.ActivePage.Shapes['ThePage'].CellsU('PrintPageOrientation')=1
        }
    }
}

<#
        .SYNOPSIS 
        Drops a shape on the page

        .DESCRIPTION
        Drops a shape (provided as a master shape) on the page.  If no X coordinate is given, the shape is positioned relative to the previous shape placed
        The shape is given a name and label.

        .PARAMETER Master
        Either the name of the master (previously registered using Register-VisioShape) or a reference to a master object.

        .PARAMETER X
        The X position used to place the shape (in inches). If this is omitted, the shape is positioned relative to the previous shape placed.

        .PARAMETER Y
        The Y position used to place the shape (in inches). 

        .PARAMETER Name
        The name for the new shape.

        .INPUTS
        None. You cannot pipe objects to Add-Extension.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        New-VisioShape MasterShapeName -Label 'My Shape' -x 5 -y 5 -Name MyShape


#>
Function New-VisioShape{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Master,$Label,$X,$Y,$Name)
    if($PSCmdlet.ShouldProcess('Visio','Drop shape on the page')){
        if($Master -is [string]){
            $Master=$script:Shapes[$Master]
        }
        if(!$Name){
            $Name=$Label
        }
 
        $p=get-VisioPage
        if($updateMode){
            $DroppedShape=$p.Shapes | Where-Object {$_.Name -eq $Label}
        }
        if(-not (get-variable DroppedShape -Scope Local -ErrorAction Ignore) -or $DroppedShape -eq $null){
            if(-not $X){
                $RelativePosition=Get-NextShapePosition
                $X=$RelativePosition.X
                $Y=$RelativePosition.Y
            }
            $DroppedShape=$p.Drop($Master.PSObject.BaseObject,$X,$Y)
            $DroppedShape.Name=$Name
        } else {
            write-verbose "Existing shape <$Label> found"
       }
        $DroppedShape.Text=$Label
        New-Variable -Name $Name -Value $DroppedShape -Scope Global -Force
        write-output $DroppedShape
        $Script:LastDroppedObject=$DroppedShape
    }

}

<#
        .SYNOPSIS 
        Draw a rectangle

        .DESCRIPTION
        Draws a rectangle on the active page with the position/size specified in the parameters

        .PARAMETER X0
        Describe The Left edge of the rectangle (in inches)

        .PARAMETER Y0
        Describe The Top edge of the rectangle (in inches)

        .PARAMETER X1
        Describe The Right edge of the rectangle (in inches)

        .PARAMETER Y1
        Describe The Bottom edge of the rectangle (in inches)

        .INPUTS
        None. You cannot pipe objects to New-VisioRectangle.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        $rect = New-VisioRectangle 1 5 2 6
        #draws a rectangle

#>
Function New-VisioRectangle{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($X0,$Y0,$X1,$Y1)
    if($PSCmdlet.ShouldProcess('Visio','Draw a rectangle on the page')){    
        $p=get-visioPage
        $p.DrawRectangle($X0,$Y0,$X1,$Y1)
    }
}

<#
        .SYNOPSIS 
        Connects two shapes

        .DESCRIPTION
        Creates a connector object between two previously drawn shapes.

        .PARAMETER From
        The shape that the connector will originate from

        .PARAMETER To
        The shape that the connector will end on

        .PARAMETER Name
        The name to assign to the connector shape

        .PARAMETER Color
        The color to draw the connector

        .PARAMETER Arrow
        Determines whether an arrow is drawn on the connector at the final end

        .PARAMETER Bidirectional
        Determines whether an arrow is drawn on the connector at the originating end

        .PARAMETER Label
        The text to be shown on the arrow


        .INPUTS
        None. You cannot pipe objects to New-VisioConnector.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        $arrow = New-VisioConnector -From WebServer -To SQLServer -name SQLConnection -Arrow -color Red -label SQL
        File.txt


#>
Function New-VisioConnector{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([array]$From,
        $To,
        $Name,
        [System.Drawing.Color]$Color,
        [switch]$Arrow,
        [switch]$bidirectional, 
    $Label)
    $ColorFormula="=rgb($($Color.R),$($Color.G),$($Color.B))"
    if($PSCmdlet.ShouldProcess('Visio','Connect shapes with a connector')){
        $CurrentPage=Get-VisioPage
        foreach($dest in $To){
            foreach($source in $From){
                if($source -is [string]){
                    $source=$CurrentPage.Shapes[$source]
                }
                if($dest -is [string]){
                    $dest=$CurrentPage.Shapes[$dest]
                }

                $CalculatedName='{0}_{1}_{2}' -f $Label,$source.Name,$dest.Name
                if($updatemode){
                    $connector=$CurrentPage.Shapes | Where-Object {$_.Name -eq $CalculatedName}
                }
                if (-not (get-variable Connector -Scope Local -ErrorAction Ignore)){
                    $source.AutoConnect($dest,0)
                    $connector=$CurrentPage.Shapes('Dynamic Connector')| Select-Object -first 1
                    $connector.Name=$CalculatedName
                }
                $connector.Text=$Label
                $connector.CellsU('LineColor').Formula=$ColorFormula
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
                Remove-variable Connector -ErrorAction SilentlyContinue
            }
        }
    }

}

<#
        .SYNOPSIS 
        Draws a container

        .DESCRIPTION
        Draws a container around previously drawn shapes on the active page.

        .PARAMETER Name
        The name to assign to the dropped shape

        .PARAMETER Conents
        A scriptblock which, when executed, outputs the objects to be contained in the container

        .PARAMETER Shape
        The master shape to use to draw the container

        .PARAMETER Label
        The text to label the container with

        .INPUTS
        None. You cannot pipe objects to New-VisioContainer.

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        New-VisioContainer -shape (Get-VisioShape Domain) -label MyDomain -contents {
            New-VisioShape -master WebServer -label BackupServer -x 5 -y 8
        }

#>
Function New-VisioContainer{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param( [string]$Name,
        [Scriptblock]$Contents,
        $Shape,
        $Label)
    if(!$Name){
        $Name=$Label
    }
    if ($shape -is [string]){
        $shape=Get-VisioShape $shape
    }
     if($PSCmdlet.ShouldProcess('Visio','Drop a container around shapes')){
        $page=Get-VisioPage
        if($Contents){
            [array]$containedObjects=& $Contents
            $firstShape=$containedObjects[0]
            if($updatemode){
                $droppedContainer=$page.Shapes | Where-Object {$_.Name -eq $Name}
            }
        
            if(get-variable droppedContainer -Scope Local -ErrorAction Ignore){
                If($droppedContainer.ContainerProperties.GetMemberShapes(16+2) -notcontains $firstShape.ID){
                    $droppedcontainer.ContainerProperties.AddMember($firstShape,2)
                }
            } else {
                $sel=New-VisioSelection $firstShape -Visible
                $droppedContainer=$page.DropContainer($Shape,$page.Application.ActiveWindow.Selection)
                $droppedContainer.Name=$Name
            } 
            $droppedContainer.ContainerProperties.SetMargin($vis.PageUnits, 0.25)
            $containedObjects | select-object -Skip 1 | % { 
                if(-not $updatemode -or ($droppedContainer.ContainerProperties.GetMemberShapes(16+2) -notcontains $_.ID)){
                    $droppedcontainer.ContainerProperties.AddMember($_,1)
                }
            }        
            $droppedContainer.ContainerProperties.FitToContents()
            $droppedContainer.Text=$Label
            $Script:LastDroppedObject=$droppedContainer

            $droppedContainer

        }
        New-Variable -Name $Name -Value $droppedContainer -Scope Global -Force
    }
}

<#
        .SYNOPSIS 
        Loads a built-in stencil  

        .DESCRIPTION
        Loads a built-in stencil  

        .PARAMETER BuiltinStencil
        Which built-in stencil to load

        .PARAMETER Name
        What name to use to reference the stencil

        .INPUTS
        None. You cannot pipe objects to Register-VisioBuiltinStencil

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioBuiltinStencil -BuiltInStencil Containers -Name VisioContainers
#>
Function Register-VisioBuiltinStencil{
    [CmdletBinding()]
    Param([ValidateSet('Backgrounds','Borders','Containers','Callouts','Legends')]
        [string]$BuiltInStencil,
    [String]$Name)
    $stencilID=@('Backgrounds','Borders','Containers','Callouts','Legends').IndexOf($BuiltInStencil)
    $stencilPath=$Visio.GetBuiltInStencilFile($stencilID,$vis.MSDefault)
    Register-VisioStencil -Path $stencilPath -Name $Name 
}
<#
        .SYNOPSIS 
        Loads a stencil and gives it a name

        .DESCRIPTION
        Loads a stencil and gives it a name

        .PARAMETER Name
        The name to use to refer to the stencil

        .PARAMETER Path
        The path to the stencil file.  Ignored with -Builtin

        .PARAMETER BuiltIn
        Flags that Path (or Name) refer to a built-in stencil.

        .INPUTS
        None. You cannot pipe objects to Register-VisioStencil.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioStencil -Name Containers -Path 'c:\temp\my containers.vssx'

        .EXAMPLE
        Register-VisioStencil -Name Connectors -Builtin

#>
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

<#
        .SYNOPSIS 
        Copies a master from a stencil and gives it a name.

        .DESCRIPTION
        Copies a master from a stencil and gives it a name.  Also creates a function with the same name to drop the shape onto the active Visio page.

        .PARAMETER Name
        The name used to refer to the shape

        .PARAMETER StencilName
        Which stencil to get the master from

        .PARAMETER MasterName
        The name of the master in the stencil

        .INPUTS
        None. You cannot pipe objects to Register-VisioShape.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioShape -Name Block -StencilName BasicShapes -MasterName Block

#>
Function Register-VisioShape{
    [CmdletBinding()]
    Param([string]$Name,
        [Alias('From')][string]$StencilName,
    [string]$MasterName)
 

    $newShape=$stencils[$StencilName].Masters | Where-Object {$_.Name -eq $MasterName}
    $script:Shapes[$Name]=$newshape
    $outerName=$Name 
    new-item -Path Function:\ -Name "global`:$outername" -value {param($Label, $X,$Y, $Name) $Shape=get-visioshape $outername; New-VisioShape $Shape $Label $X $Y -name $Name}.GetNewClosure() -force  | out-null

}
<#
        Copies a master for a container from a stencil and gives it a name.

        .DESCRIPTION
        Copies a master for a container from a stencil and gives it a name.  Also creates a function with the same name to drop the container (with contents) onto the active Visio page.

        .PARAMETER Name
        The name used to refer to the shape

        .PARAMETER StencilName
        Which stencil to get the master from

        .PARAMETER MasterName
        The name of the master in the stencil

        .INPUTS
        None. You cannot pipe objects to Register-VisioContainer.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioContainer -Name BasicContainer -StencilName Containers -MasterName Plain

#>
Function Register-VisioContainer{
    [CmdletBinding()]
    Param([string]$Name,
        [Alias('From')][string]$StencilName,
    [string]$MasterName)
 

    $newShape=$stencils[$StencilName].Masters | Where-Object {$_.Name -eq $MasterName}
    $script:Shapes[$Name]=$newshape
    $outerName=$Name
    new-item -Path Function:\ -Name "global`:$outername" -value {param($Label,$Contents,$Name) New-VisioContainer -label $Label -contents $Contents -shape $outername -name $Name}.GetNewClosure() -force  | out-null

}
<#
        .SYNOPSIS 
        Saves a "nickname" for a certain style of connector

        .DESCRIPTION
        Saves a "nickname" for a certain style of connector

        .PARAMETER Name
        The name to use to refer to this style of connector

        .PARAMETER Color
        The color to draw the connector in

        .PARAMETER Arrow
        Whether to put an arrow at the end of the connector

        .PARAMETER Bidirectional
        Whether to put an arrow at the beginning of the connector


        .INPUTS
        None. You cannot pipe objects to Register-VisioConnector.

        .OUTPUTS
        None

        .EXAMPLE
        Register-VisioConnector -Name HTTP -Color Black -Ar
#>
Function Register-VisioConnector{
    [CmdletBinding()]
    Param([string]$Name,
        [System.Drawing.Color]$Color,
        [switch]$Arrow,
    [switch]$bidirectional)
    new-item -Path Function:\ -Name "global`:$Name" -value {param($From,$To,$Label) New-VisioConnector -from $From -to $To -name $Name -color $Color -Arrow:$Arrow.IsPresent -bidirectional:$bidirectional.IsPresent $Label}.GetNewClosure() -force  | out-null
}


<#
        .SYNOPSIS 
        Retrieves a saved shape definition

        .DESCRIPTION
        Retrieves a saved shape definition

        .PARAMETER Name
        Describe Parameter1

        .INPUTS
        None. You cannot pipe objects to Get-VisioShape

        .OUTPUTS
        Visio.Shape

        .EXAMPLE
        Get-VisioShape Block

#>
Function Get-VisioShape{
    [CmdletBinding()]
    Param([string]$Name)
    $script:Shapes[$Name]
}

<#
        .SYNOPSIS 
        Sets the hyperlink on a shape to the given address.

        .DESCRIPTION
        Sets the hyperlink on a shape to the given address.

        .PARAMETER Shape
        The shape you want the hyperlink on

        .PARAMETER Link
        The address of the hyperlink

        .INPUTS
        None. You cannot pipe objects to New-VisioHyperlink.

        .OUTPUTS
        None

        .EXAMPLE
        New-VisioHyperlink -shape $rectangle -link http://google.com
        File.txt

#>
Function New-VisioHyperlink{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Shape,
    $Link)
    if($PSCmdlet.ShouldProcess('Visio','Create a hyperlink on a shape')){
        $CurrentPage=Get-VisioPage
        if($Shape -is [string]){
            $Shape=$CurrentPage.Shapes[$Shape]
        }
        $LinkObject=$Shape.AddHyperLink()
        $LinkObject.Address=$Link
    }
}

<#
        .SYNOPSIS 
        Creates a selection object using the given shapes.

        .DESCRIPTION
        Creates a selection object using the given shapes.  If -Visible is passed, the selection is shown in the application.

        .PARAMETER Objects
        The objects to be selected

        .PARAMETER Visible
        Whether the selection is visible in the application

        .INPUTS
        What can be piped in
        None. You cannot pipe objects to New-VisioSelection.

        .OUTPUTS
        Visio.Selection

        .EXAMPLE
        New-VisioSelection -Objects Server1,Server2
#>
Function New-VisioSelection{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([array]$Objects,[switch]$Visible)
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


<#
        .SYNOPSIS 
        Sets the value of a shape data field.

        .DESCRIPTION
        Sets the value of a shape data field.

        .PARAMETER Shape
        The shape that has the shape data

        .PARAMETER Name
        The name of the shape data field to set

        .PARAMETER Value
        The value to set the shape data to

        .INPUTS
        None. You cannot pipe objects to Set-VisioShapeData.

        .OUTPUTS
        None

        .EXAMPLE
        Set-VisioShapeData -shape $WebServer -Name IPAddress -Value 10.1.1.5
#>
Function Set-VisioShapeData{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param($Shape,
        $Name,
    $Value)
        if($PSCmdlet.ShouldProcess('Visio','Set a value for a custom shape data element')){
        $Shape.Cells("Prop.$Name").Formula="=`"$value`""
    }
}

<#
        .SYNOPSIS 
        Returns a shape data field from a shape

        .DESCRIPTION
        Returns a shape data field from a shape

        .PARAMETER Shape
        The shape that has the shape data

        .PARAMETER Name
        Which shape data field you want the value from

        .INPUTS
        None. You cannot pipe objects to Get-VisioShapeData.

        .OUTPUTS
        String

        .EXAMPLE
        Get-VisioShapeData -shape $webServer -Name IPAddress
#>
Function Get-VisioShapeData{
    [CmdletBinding()]
    Param($Shape,
    $Name)
    if($PSCmdlet.ShouldProcess('Visio','Retrieve the value from a custom shape data element')){
        $Shape.Cells("Prop.$Name").Formula.TrimStart('"').TrimEnd('"') 
    }
}

<#
        .SYNOPSIS 
        Saves the diagram and optionally exits Visio

        .DESCRIPTION
        Saves the diagram and optionally exits Visio

        .PARAMETER Close
        Whether to exit Visio or not

        .INPUTS
        None. You cannot pipe objects to Complete-Diagram.

        .OUTPUTS
        None

        .EXAMPLE
        Complete-Diagram
#>
Function Complete-Diagram{
    [CmdletBinding()]
    Param([switch]$Close)
    $script:updateMode=$false
    $Visio.ActiveDocument.Save() 
    if($Close){
        $Visio.Quit()
    }
}

<#
        .SYNOPSIS 
        Creates a new Visio Layer and adds the given objects to it.

        .DESCRIPTION
        Long DescriptionCreates a new Visio Layer and adds the given objects to it.

        .PARAMETER LayerName
        The name for the new layer

        .PARAMETER Contents
        The objects to be included in the layer

        .PARAMETER Preserve
        Whether to preserve the existing layer assignments for these objecfts.

        .INPUTS
        None. You cannot pipe objects to New-VisioLayer.

        .OUTPUTS
        Visio.Layer

        .EXAMPLE
        New-VisioLayer -Layer WebServers -Contents WebServer1,WebServer2 -Preserve
#>
Function New-VisioLayer{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([string]$LayerName,[Array]$Contents,[switch]$Preserve)
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
        if ($Contents -is [scriptblock]){
            $Contents = & $Contents
        }
        foreach($item in [array]$Contents){
            if($item -is [string]){
                $item=$p.Shapes[$item] 
            }
            $layer.Add($item,$AddOption)
        }
    }
}


<#
        .SYNOPSIS 
        Sets the next position to place a shape using relative positioning

        .DESCRIPTION
        Sets the next position to place a shape using relative positioning

        .INPUTS
        None.  

        .OUTPUTS
        None

        .EXAMPLE
        Set-NextShapePosition -x 5 -y 6
         

#>
Function Set-NextShapePosition{
    Param($X,$Y)
    $script:LastDroppedObject=@{X=$X;y=$Y}
}

<#
        .SYNOPSIS 
        Returns the next position to place a shape using relative positioning

        .DESCRIPTION
        Returns the next position to place a shape using relative positioning

        .INPUTS
        None. You cannot pipe objects to Get-NextShapePosition.

        .OUTPUTS
        HashTable

        .EXAMPLE
        Get-NextShapePosition  
        #returns a hashtable with X and Y position of next shape to place.

#>
Function Get-NextShapePosition{
    [CmdletBinding()]
    Param()
    if($LastDroppedObject -eq 0){
        #nothing dropped yet, start at top-left-ish
        $p=Get-VisioPage
        
        return @{X=1;Y=$p.Pagesheet.Cells('PAgeHeight').ResultIU-1}
    } elseif ($LastDroppedObject -is [hashtable]) {
        return $LastDroppedObject
    } else {
        if($RelativeOrientation -eq 'Horizontal'){
            $X=$LastDroppedObject.Cells('PinX').ResultIU + $LastDroppedObject.Cells('Width').ResultIU + 0.25
            $Y=$LastDroppedObject.Cells('PinY').ResultIU 
        } else {
            $X=$LastDroppedObject.Cells('PinX').ResultIU 
            $Y=$LastDroppedObject.Cells('PinY').ResultIU - $LastDroppedObject.Cells('Height').ResultIU - 0.25
        }
        Return @{X=$X;Y=$Y}
    }
}

<#
        .SYNOPSIS 
        Changes the direction VisioBot3000 uses when placing shapes using relative positioning

        .DESCRIPTION
        Changes the direction VisioBot3000 uses when placing shapes using relative positioning

        .PARAMETER Orientation
        Either vertical or Horizontal

        .INPUTS
        None. You cannot pipe objects to Set-RelativePositionDirection.

        .OUTPUTS
        None

        .EXAMPLE
        Set-RelativePositionDirection Horizontal
        File.txt

#>
Function Set-RelativePositionDirection{
    [CmdletBinding(SupportsShouldProcess=$True)]
    Param([ValidateSet('Horizontal','Vertical')]$Orientation)
    if($PSCmdlet.ShouldProcess('Visio','Sets VisioBot''s orientation for relative positioning')){
        $script:RelativeOrientation=$Orientation
    }
}

<#
        .SYNOPSIS 
        Sets the text of multiple objects that have already been created
        .DESCRIPTION
        Sets the text of multiple objects that have already been created. Useful for updating static objects that are part of a template, such as title, author, etc.
        .PARAMETER Map
        A hashtable mapping names of objects to the text you want them to have.
        .INPUTS
        None. You cannot pipe objects to Set-VisioText.
        .OUTPUTS
        None
        .EXAMPLE
        Set-VisioText -Map @{Title='My first Visio Digram';
        Author='Mike';
        CreatedOn="$(get-date)"}
#> 

function Set-VisioText{
    [CmdletBinding()]
    Param([Hashtable]$Map)
    
    foreach($key in $Map.Keys){
        $text=$map[$key]
        $p=Get-VisioPage
        while($key.Contains('/')){
            $prefix,$key=$key.split('/',2)
            $p=$p.Shapes[$prefix]
        }
        $Shape=$p.Shapes[$key]
        $Shape.Characters.Text="$text"
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
New-Alias -Name Legend -value Set-VisioText