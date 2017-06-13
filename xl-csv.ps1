function xl-csv
{
#This function copies input Excel file as a csv with the same name in the same directory
#The contents of the CSV will be stored as the variable $CSV
 
 [cmdletbinding()]
 param
 (
  [parameter(Mandatory=$true)]
  [string] $XL_Path
 )
 
 $CSV=$XL_Path|%{$_ -replace "(\s*)\.\w*","$1.csv"}
 $Excel=New-Object -ComObject excel.application
 $Excel.visible=$false
 $Excel.displayalerts=$false
 $wb=$excel.workbooks.open($XL_Path)
 $wb.saveas($CSV,6)
 $Excel.quit()
 $CSV=import-csv $CSV

}

