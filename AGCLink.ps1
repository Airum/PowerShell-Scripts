#Given a directory, this script finds the two AGC output files that are active, then compares their values
#This output is redirected  to an Excel file, and a VBA macro is activated to present it nicely to an auditor

Param
    (
        [Parameter(Mandatory=$true)] 
        [string] $WorkingDirectory
    )

#Find which cases were active to compare
Get-ChildItem ($WorkingDirectory +"\*.csv")|where{$_.name -match "CC_"}|%{if(gc $_| select-string -pattern  "VALUTIME='error'" ){remove-item $_}}
$list=Get-ChildItem ($WorkingDirectory +"\*.csv")|where{$_.name -match "CC_"}
$file1='\'+$list.name.Get(0)
$file2='\'+$list.name.Get(1)

#Give the systems meaningful names
if($file1 -match "DCC"){$name1="D"}
else{$name1="M"}
if($file1 -match "VMS"){$name1+="_2.1"}
else{$name1+="_2.6"}
if($file1 -match "CTL"){$name1+="_Control"}
else{$name1+="_Backup"}

if($file2 -match "DCC"){$name2="D"}
else{$name2="M"}
if($file2 -match "VMS"){$name2+="_2.1"}
else{$name2+="_2.6"}
if($file2 -match "CTL"){$name2+="_Control"}
else{$name2+="_Backup"}

#Give both files identical headers so they can be compared in PS, also adding the 'set #' column
$sys1=import-csv ($WorkingDirectory+$file1) -Header 'Type',	$name1	,"ID=",	'CTLSTATE=',	"LASTMSG=",	"LINKUP=",	"LOCAL=",	"REMOTE=",	"ALLOWCTL=T",	"TITLE=",'ONLINE=T',	'USRFLG01=T',	'USRFLG02=T',	'USRFLG03=F',	'USRFLG04=F',	'USRFLG05=F',	'USRFLG06=T',	'USRFLG07=F',	'USRFLG08=T',	'USRFLG09=T',	'USRFLG10=F',	'USRFLG11=F',	'USRFLG12=F',	'USRFLG13=F',	'USRFLG14=T',	'USRFLG15=F',	'USRFLG16=F',	'USRFLG17=F',	'USRFLG18=F',	'USRFLG19=F',	'USRFLG20=F',	'USRFLG21=F',	'USRFLG22=F',	'USRFLG23=F',	'USRFLG24=F',	'USRFLG25=F',	'USRFLG26=F',	'USRFLG27=F',	'USRFLG28=F',	'USRFLG29=F',	'USRFLG30=F',	'USRFLG31=F',	'USRFLG32=F',	"NETIOSYS='GD'",	"REMMODE='OPER'", 'set #'|where{$_.'Type' -ne 'system'}
$sys2=import-csv ($WorkingDirectory+$file2) -Header 'Type',	$name2	,"ID=",	'CTLSTATE=',	"LASTMSG=",	"LINKUP=",	"LOCAL=",	"REMOTE=",	"ALLOWCTL=T",	"TITLE=",'ONLINE=T',	'USRFLG01=T',	'USRFLG02=T',	'USRFLG03=F',	'USRFLG04=F',	'USRFLG05=F',	'USRFLG06=T',	'USRFLG07=F',	'USRFLG08=T',	'USRFLG09=T',	'USRFLG10=F',	'USRFLG11=F',	'USRFLG12=F',	'USRFLG13=F',	'USRFLG14=T',	'USRFLG15=F',	'USRFLG16=F',	'USRFLG17=F',	'USRFLG18=F',	'USRFLG19=F',	'USRFLG20=F',	'USRFLG21=F',	'USRFLG22=F',	'USRFLG23=F',	'USRFLG24=F',	'USRFLG25=F',	'USRFLG26=F',	'USRFLG27=F',	'USRFLG28=F',	'USRFLG29=F',	'USRFLG30=F',	'USRFLG31=F',	'USRFLG32=F',	"NETIOSYS='GD'",	"REMMODE='OPER'", 'set #'|where{$_.'Type' -ne 'system'}	

#Fill in the 'set #' column, linking values to points.  Separate agcpnt from pntval.
$i=0
$sys1|%{if($_.'Type' -eq 'AGCPNT'){($i=$i+1); $_.'set #'=$i} else{$_.'set #'=$i}}>$null
$sys1agc=$sys1|where{$_.'Type' -eq 'AGCPNT'}
$sys1val=$sys1|where{$_.'Type' -eq 'pntval'}


$i=0
$sys2|%{if($_.'Type' -eq 'AGCPNT'){($i=$i+1); $_.'set #'=$i} else{$_.'set #'=$i}}>$null
$sys2agc=$sys2|where{$_.'Type' -eq 'AGCPNT'}
$sys2val=$sys2|where{$_.'Type' -eq 'pntval'} 

#Compare I,L,T,& RVALUE between both systems' pntval .CSVs. 
$compval=compare-Object $sys1val $sys2val -Property 'ctlstate=','LINKUP=','LASTMSG=','LOCAL=' -SyncWindow 0 -PassThru
$compval1=$compval |where{$_.sideindicator -eq "=>"}
$compval2=$compval |where{$_.sideindicator -eq "<="}



#Recombine pntval and agcpnt csv's, only keeping pntvals that differed and the agcpnts that correspond with those
$listval1=$compval1|Select-Object -ExpandProperty "set #"
$changes1=($sys1agc|where{ $listval1 -contains $_.'set #'})+$compval1|sort 'set #', 'Type' |Export-Csv ($WorkingDirectory+"\Discrepancies1.csv")  -NoTypeInformation

$listval2=$compval2|Select-Object -ExpandProperty "set #"
$changes2=($sys2agc|where{ $listval2 -contains $_.'set #'})+$compval2|sort 'set #', 'Type' |Export-Csv ($WorkingDirectory+"\Discrepancies2.csv")  -NoTypeInformation 


#Run Excel Macros to highlight and display errors
$excel= New-Object -ComObject "excel.application"
$macro=$excel.Workbooks.open($workingDirectory+"\Highlighter.xlsm")
$excel.run("Sides")
$excel.Visible=$true
