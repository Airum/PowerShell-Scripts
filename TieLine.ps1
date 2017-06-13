Param
    (
        [Parameter(Mandatory=$true)] 
        [string] $WorkingDirectory,

        [Parameter(Mandatory=$true)]
        [string] $BPA,

        [Parameter(Mandatory = $true)]
        [string] $Peak
        )

#Set up some files in the same directory to be used later
$PeakN=$WorkingDirectory+"PeakN.csv"
$BPAN=$WorkingDirectory+"BPAN.csv"
$BPAS=$WorkingDirectory+"BPAS.csv"
$temp=$WorkingDirectory+"temp.txt"

#Make inputs CSV
(gc $BPA).Trim()| %{$_ -replace "\s+",","}|Set-Content $BPAN,$BPAS
(gc $Peak).Trim()| %{$_ -replace "\s+",","}|Set-content $PeakN

#Declare variables for adding headers
$a= "Num,Op_Area,Line_Name,Meter_St,From_St,From_Nd,To_St,To_Nd,Connect_Area"
$b= "Num,Op_Area,Line_Name,Meter_St,To_St,To_Nd,From_St,From_Nd,Connect_Area"

#Add headers to BPA and Peak files
New-Item $temp -ItemType file
Set-Content $temp -Value $a
Add-Content $temp -Value (gc $BPAN)
Remove-Item $BPAN
Rename-Item $temp -NewName $BPAN

New-Item $temp -ItemType file
Set-Content $temp -Value $a
Add-Content $temp -Value (gc $PeakN)
remove-item $peakN
Rename-Item $temp -NewName $PeakN

#Switch From and To headers in BPA
New-Item $temp -ItemType file
Set-Content $temp -Value $b
Add-Content $temp -Value (gc $BPAS)
Remove-Item $BPAS
Rename-Item $temp -NewName $BPAS

#Find all Line Names that match Up perfectly
$perfect=compare-object (import-csv $bpaN) (import-csv $peakN) -Property "Op_Area","Line_Name","From_St","From_Nd","To_St","To_Nd" -includeequal -excludedifferent|Select-Object -ExpandProperty line_name

#Remove them from both lists, using the switched bpa data
$peakdata= (import-csv $peakN)|where{$perfect -notcontains $_.line_name}
$BPAdata= (import-csv $BPAs)|where{$perfect -notcontains $_.line_name}

#Find all Line Names that still don't match after reversal
$perfect2=compare-object $bpadata $peakdata -Property "Op_Area","Line_Name","From_St","From_Nd","To_St","To_Nd" -PassThru

#Change arrows to more specific info
$perfect2|% { 
      if ($_.sideindicator -eq '<=')
        {$_.sideindicator = 'BPA ONLY'}

      if ($_.sideindicator -eq '=>')
        {$_.sideindicator = 'Peak ONLY'}
     }

#If errors were found, output a report, otherwise give success message to user
if ($perfect2.length -eq 0) {write-host "No Discrepancies Found!" -ForegroundColor Red} Else {$perfect2|Export-Csv ($WorkingDirectory + "Discrepancies.csv")}

#Cleanup
Remove-Item $BPAN, $BPAS, $PeakN
