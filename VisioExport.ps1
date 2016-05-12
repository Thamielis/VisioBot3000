function Convert-VisioObjectToPSObject{
[CmdletBinding()]
Param([Parameter(ValueFromPipeline=$true)]$object) 
    
    $objectHash=@{Name=$object.Name;
                      Text=$object.Text}
    if(($object | get-member Master) -and $object.Master){
       $objectHash.Add('Type',$object.Master.Name)
    }
    if($object | get-member PageSheet){
        #get the shapes that aren't in containers
        $containedObjects= $object.Shapes | Where-Object {$_.MemberOfContainers.Count -eq 0} |foreach {Export-VisioObject $_}
    } elseif($object | get-member ContainerProperties){
        #get the top-level objects contained
        if($object.ContainerProperties){
            $containedObjectIDs=$object.ContainerProperties.GetMemberShapes(16+2) 
            $containedObjects=$ContainedObjectIDs|foreach {Export-VisioObject ($object.ContainingPage.Shapes | WHERE id -EQ $_)}
        }
    }
    if($containedObjects){
        $objectHash.Add('Contents',$containedObjects)
    }
    [PSCustomObject]$objectHash

}