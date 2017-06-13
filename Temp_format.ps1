#This script creates the input for substations to be added to the dynamic temperature app.
 #Input the segment list and the line limits csv
 
 function Join-Object{ 
 [CmdletBinding()]
    Param
    (
        [Parameter(Mandatory=$true)]
        [object[]] $Left,

        # List to join with $Left
        [Parameter(Mandatory=$true)]
        [object[]] $Right,

        [Parameter(Mandatory = $true)]
        [string] $LeftJoinProperty,

        [Parameter(Mandatory = $true)]
        [string] $RightJoinProperty,

        [object[]]$LeftProperties = '*',

        # Properties from $Right we want in the output.
        # Like LeftProperties, each can be a plain name, wildcard or hashtable. See the LeftProperties comments.
        [object[]]$RightProperties = '*',

        [validateset( 'AllInLeft', 'OnlyIfInBoth', 'AllInBoth', 'AllInRight')]
        [Parameter(Mandatory=$false)]
        [string]$Type = 'AllInLeft',

        [string]$Prefix,
        [string]$Suffix
    )
    Begin
    {
        function AddItemProperties($item, $properties, $hash)
        {
            if ($null -eq $item)
            {
                return
            }

            foreach($property in $properties)
            {
                $propertyHash = $property -as [hashtable]
                if($null -ne $propertyHash)
                {
                    $hashName = $propertyHash["name"] -as [string]         
                    $expression = $propertyHash["expression"] -as [scriptblock]

                    $expressionValue = $expression.Invoke($item)[0]
            
                    $hash[$hashName] = $expressionValue
                }
                else
                {
                    foreach($itemProperty in $item.psobject.Properties)
                    {
                        if ($itemProperty.Name -like $property)
                        {
                            $hash[$itemProperty.Name] = $itemProperty.Value
                        }
                    }
                }
            }
        }

        function TranslateProperties
        {
            [cmdletbinding()]
            param(
                [object[]]$Properties,
                [psobject]$RealObject,
                [string]$Side)

            foreach($Prop in $Properties)
            {
                $propertyHash = $Prop -as [hashtable]
                if($null -ne $propertyHash)
                {
                    $hashName = $propertyHash["name"] -as [string]         
                    $expression = $propertyHash["expression"] -as [scriptblock]

                    $ScriptString = $expression.tostring()
                    if($ScriptString -notmatch 'param\(')
                    {
                        Write-Verbose "Property '$HashName'`: Adding param(`$_) to scriptblock '$ScriptString'"
                        $Expression = [ScriptBlock]::Create("param(`$_)`n $ScriptString")
                    }
                
                    $Output = @{Name =$HashName; Expression = $Expression }
                    Write-Verbose "Found $Side property hash with name $($Output.Name), expression:`n$($Output.Expression | out-string)"
                    $Output
                }
                else
                {
                    foreach($ThisProp in $RealObject.psobject.Properties)
                    {
                        if ($ThisProp.Name -like $Prop)
                        {
                            Write-Verbose "Found $Side property '$($ThisProp.Name)'"
                            $ThisProp.Name
                        }
                    }
                }
            }
        }

        function WriteJoinObjectOutput($leftItem, $rightItem, $leftProperties, $rightProperties)
        {
            $properties = @{}

            AddItemProperties $leftItem $leftProperties $properties
            AddItemProperties $rightItem $rightProperties $properties

            New-Object psobject -Property $properties
        }

        #Translate variations on calculated properties.  Doing this once shouldn't affect perf too much.
        foreach($Prop in @($LeftProperties + $RightProperties))
        {
            if($Prop -as [hashtable])
            {
                foreach($variation in ('n','label','l'))
                {
                    if(-not $Prop.ContainsKey('Name') )
                    {
                        if($Prop.ContainsKey($variation) )
                        {
                            $Prop.Add('Name',$Prop[$Variation])
                        }
                    }
                }
                if(-not $Prop.ContainsKey('Name') -or $Prop['Name'] -like $null )
                {
                    Throw "Property is missing a name`n. This should be in calculated property format, with a Name and an Expression:`n@{Name='Something';Expression={`$_.Something}}`nAffected property:`n$($Prop | out-string)"
                }


                if(-not $Prop.ContainsKey('Expression') )
                {
                    if($Prop.ContainsKey('E') )
                    {
                        $Prop.Add('Expression',$Prop['E'])
                    }
                }
            
                if(-not $Prop.ContainsKey('Expression') -or $Prop['Expression'] -like $null )
                {
                    Throw "Property is missing an expression`n. This should be in calculated property format, with a Name and an Expression:`n@{Name='Something';Expression={`$_.Something}}`nAffected property:`n$($Prop | out-string)"
                }
            }        
        }

        $leftHash = @{}
        $rightHash = @{}

        # Hashtable keys can't be null; we'll use any old object reference as a placeholder if needed.
        $nullKey = New-Object psobject
        
        $bound = $PSBoundParameters.keys -contains "InputObject"
        if(-not $bound)
        {
            [System.Collections.ArrayList]$LeftData = @()
        }
    }
    Process
    {
        #We pull all the data for comparison later, no streaming
        if($bound)
        {
            $LeftData = $Left
        }
        Else
        {
            foreach($Object in $Left)
            {
                [void]$LeftData.add($Object)
            }
        }
    }
    End
    {
        foreach ($item in $Right)
        {
            $key = $item.$RightJoinProperty

            if ($null -eq $key)
            {
                $key = $nullKey
            }

            $bucket = $rightHash[$key]

            if ($null -eq $bucket)
            {
                $bucket = New-Object System.Collections.ArrayList
                $rightHash.Add($key, $bucket)
            }

            $null = $bucket.Add($item)
        }

        foreach ($item in $LeftData)
        {
            $key = $item.$LeftJoinProperty

            if ($null -eq $key)
            {
                $key = $nullKey
            }

            $bucket = $leftHash[$key]

            if ($null -eq $bucket)
            {
                $bucket = New-Object System.Collections.ArrayList
                $leftHash.Add($key, $bucket)
            }

            $null = $bucket.Add($item)
        }

        $LeftProperties = TranslateProperties -Properties $LeftProperties -Side 'Left' -RealObject $LeftData[0]
        $RightProperties = TranslateProperties -Properties $RightProperties -Side 'Right' -RealObject $Right[0]

        #I prefer ordered output. Left properties first.
        [string[]]$AllProps = $LeftProperties

        #Handle prefixes, suffixes, and building AllProps with Name only
        $RightProperties = foreach($RightProp in $RightProperties)
        {
            if(-not ($RightProp -as [Hashtable]))
            {
                Write-Verbose "Transforming property $RightProp to $Prefix$RightProp$Suffix"
                @{
                    Name="$Prefix$RightProp$Suffix"
                    Expression=[scriptblock]::create("param(`$_) `$_.'$RightProp'")
                }
                $AllProps += "$Prefix$RightProp$Suffix"
            }
            else
            {
                Write-Verbose "Skipping transformation of calculated property with name $($RightProp.Name), expression:`n$($RightProp.Expression | out-string)"
                $AllProps += [string]$RightProp["Name"]
                $RightProp
            }
        }

        $AllProps = $AllProps | Select -Unique

        Write-Verbose "Combined set of properties: $($AllProps -join ', ')"

        foreach ( $entry in $leftHash.GetEnumerator() )
        {
            $key = $entry.Key
            $leftBucket = $entry.Value

            $rightBucket = $rightHash[$key]

            if ($null -eq $rightBucket)
            {
                if ($Type -eq 'AllInLeft' -or $Type -eq 'AllInBoth')
                {
                    foreach ($leftItem in $leftBucket)
                    {
                        WriteJoinObjectOutput $leftItem $null $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
            else
            {
                foreach ($leftItem in $leftBucket)
                {
                    foreach ($rightItem in $rightBucket)
                    {
                        WriteJoinObjectOutput $leftItem $rightItem $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
        }

        if ($Type -eq 'AllInRight' -or $Type -eq 'AllInBoth')
        {
            foreach ($entry in $rightHash.GetEnumerator())
            {
                $key = $entry.Key
                $rightBucket = $entry.Value

                $leftBucket = $leftHash[$key]

                if ($null -eq $leftBucket)
                {
                    foreach ($rightItem in $rightBucket)
                    {
                        WriteJoinObjectOutput $null $rightItem $LeftProperties $RightProperties | Select $AllProps
                    }
                }
            }
        }
    }
}
Param
    (
        [Parameter(Mandatory=$true)]
        [string] $Limits,

        [Parameter(Mandatory=$true)]
        [string] $Segs,

        [Parameter(Mandatory=$true)]
        [string] $New,		
		
        [Parameter(Mandatory=$true)]
        [string] $Out

     )


#Turn segments and lines into usable csvs and import, filter down to BPA segments only
$Ratings=gc $Limits|%{$_ -replace "RM,DURATION,EMERGENCY,DURATION,LOADSHEDDING,DURATION","RM,DURATION1,EMERGENCY,DURATION2,LOADSHEDDING,DURATION3"}|ConvertFrom-Csv
if(Test-Path($New)){Remove-Item $New}
New-Item $New -ItemType file
Add-Content $New "Type,Name,CO,ID"
Add-Content $New (gc($segs))
$Segs=Import-Csv $New #|where{$_.CO -eq "'BPA'"}

#Make Line ID on ratings and Segs match
$Ratings|%{$_.Line_ID=$_.Line_id+"_"+$_.LN_ID}
$Segs|%{$_.name=$_.name -replace "'",""}

#Remove duplicate line names (Tie Lines)
$Segs2=$Segs|sort -Property "Name" -Unique

#Join lists based on line names, filter out ones that don't match
$TempRtg= Join-Object $Segs2 $Ratings -LeftJoinProperty Name -RightJoinProperty Line_id -LeftProperties "ID", "CO" -RightProperties "Line_ID","Temperature_Index","NORM","EMERGENCY","LOADSHEDDING"|where{$_.line_id -match "_"}

#Format, code based on temperature index
$i=1
$TempRtg2=$TempRtg|%
{
 switch($_.Temperature_index)
 {
  14 {$i=1}
  32 {$i=2}
  41 {$i=3}
  50 {$i=4}
  68 {$i=5}
  86 {$i=6}
  104{$i=7}
 };

 if($_.co -eq "'PSEI'")
 {
  $_="pos seg="+$_.ID+"; /RTG1SCAL=1; /RTG2SCAL=1; /RTG3SCAL=1; /I__TEMPMDL=1; /TEMPMDL=TEMPMDL1; /entmp=f; insert rating; /ID="+$_.Temperature_index+"; /I__TEMPPT="+$i+";/P__SEG="+$_.ID+"; /RTG1="+$_.NORM+"; /RTG2="+$_.Emergency+"; /RTG3="+$_.Loadshedding+";";$_
 } 
 if($_.co -eq "'BPA'")
 {
  $_="pos seg="+$_.ID+"; /RTG1SCAL=1; /RTG2SCAL=1; /RTG3SCAL=1; /I__TEMPMDL=2; /TEMPMDL=TEMPMDL2; /entmp=t; insert rating; /ID="+$_.Temperature_index+"; /I__TEMPPT="+$i+";/P__SEG="+$_.ID+"; /RTG1="+$_.NORM+"; /RTG2="+$_.Emergency+"; /RTG3="+$_.Loadshedding+";";$_
 } 
 if($_.co -eq "'TPWR'")
 {
  $_="pos seg="+$_.ID+"; /RTG1SCAL=1; /RTG2SCAL=1; /RTG3SCAL=1; /I__TEMPMDL=3; /TEMPMDL=TEMPMDL3; /entmp=f; insert rating; /ID="+$_.Temperature_index+"; /I__TEMPPT="+$i+";/P__SEG="+$_.ID+"; /RTG1="+$_.NORM+"; /RTG2="+$_.Emergency+"; /RTG3="+$_.Loadshedding+";";$_
 } 
 if($_.co -eq "'PNM'")
 {
  $_="pos seg="+$_.ID+"; /RTG1SCAL=1; /RTG2SCAL=1; /RTG3SCAL=1; /I__TEMPMDL=3; /TEMPMDL=TEMPMDL3; /entmp=f; insert rating; /ID="+$_.Temperature_index+"; /I__TEMPPT="+$i+";/P__SEG="+$_.ID+"; /RTG1="+$_.NORM+"; /RTG2="+$_.Emergency+"; /RTG3="+$_.Loadshedding+";";$_
 } 
 else
 {
  $_="pos seg="+$_.ID+"; /RTG1SCAL=1; /RTG2SCAL=1; /RTG3SCAL=1; insert rating; /ID="+$_.Temperature_index+"; /I__TEMPPT="+$i+";/P__SEG="+$_.ID+"; /RTG1="+$_.NORM+"; /RTG2="+$_.Emergency+"; /RTG3="+$_.Loadshedding+";";$_}
 }

#Reverse order so high temperatures are first
[array]::Reverse($TempRtg2)

#Output, cleanup old outputs
if(test-path $Out){remove-item $Out}
$TempRtg2|Out-File $Out -Encoding ascii

