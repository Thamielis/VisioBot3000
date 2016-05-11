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