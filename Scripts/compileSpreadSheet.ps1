param($statsLoc)
# Excel Constants
 
# MsoTriState
$loc = $statsLoc
$currentDate = (get-date).ToString("yyyyMMdd")
$csvLoc = $statsLoc + "Hosts.csv"
$data = import-csv $csvLoc
$servers = $data | select -ExpandProperty HostName -Unique | sort

#loop each server to create Excel sheet and tie data to it.
foreach ($server in $servers){
	Set-Variable msoFalse 0 -Option Constant -ErrorAction SilentlyContinue 
	Set-Variable msoTrue 1 -Option Constant -ErrorAction SilentlyContinue
	 
	Set-Variable cellWidth 48 -Option Constant -ErrorAction SilentlyContinue
	Set-Variable cellHeight 15 -Option Constant -ErrorAction SilentlyContinue
	
	if (Test-Path  ($loc + 'Images\Host\' + $server + '_' + $currentDate + '.png'))
	{ 
		$points = $data | Where-Object {($_.HostName -eq $server)}
		$xl = New-Object -ComObject Excel.Application -Property @{
			Visible = $true
			DisplayAlerts = $false
		}
		 
		$wb = $xl.WorkBooks.Add()
		$sh = $wb.Sheets.Item('Sheet1')
		$rowStart = 3
	 
		#Set Headers
		$sh.Cells.Item(2,1)= "HostName"
		$sh.Cells.Item(2,1).Font.Bold= $true
		$sh.Cells.Item(2,2)= "MemMax"
		$sh.Cells.Item(2,2).Font.Bold= $true
		$sh.Cells.Item(2,3)= "MemAvg"
		$sh.Cells.Item(2,3).Font.Bold= $true
		$sh.Cells.Item(2,4)= "MemMin"
		$sh.Cells.Item(2,4).Font.Bold= $true
		$sh.Cells.Item(2,5)= "CPUMax"
		$sh.Cells.Item(2,5).Font.Bold= $true
		$sh.Cells.Item(2,6)= "CPUAvg"
		$sh.Cells.Item(2,6).Font.Bold= $true
		$sh.Cells.Item(2,7)= "CPUMin"
		$sh.Cells.Item(2,7).Font.Bold= $true
		$sh.Cells.Item(2,8)= "Date"
		$sh.Cells.Item(2,8).Font.Bold= $true
		
		#loop excel data from hosts.csv
		foreach ($point in $points){
			$sh.Cells.Item($rowStart,1)= $point.HostName + "%"
			$sh.Cells.Item($rowStart,2)= $point.MemMax + "%"
			$sh.Cells.Item($rowStart,3)= $point.MemAvg + "%"
			$sh.Cells.Item($rowStart,4)= $point.MemMin + "%"
			$sh.Cells.Item($rowStart,5)= $point.CPUMax + "%"
			$sh.Cells.Item($rowStart,6)= $point.CPUAvg + "%"
			$sh.Cells.Item($rowStart,7)= $point.CPUMin + "%"
			$sh.Cells.Item($rowStart,8)= $point.Date
			$rowStart = $rowStart + 1
		}
		
		# arguments to insert the image through the Shapes.AddPicture Method
	
		$imgPath = $loc + 'Images\Host\' + $server + '_' + $currentDate + '.png'
		$LinkToFile = $msoFalse
		$SaveWithDocument = $msoTrue
		$Left = $cellWidth * 12
		$Top = $cellHeight * 2
		$Width = $cellWidth * 18
		$Height = $cellHeight * 28
		 
		# add image to the Sheet
		 
		$img = $sh.Shapes.AddPicture($imgPath, $LinkToFile, $SaveWithDocument,$Left, $Top, $Width, $Height)
		#$xl.Speech.Speak('Add an image to the Sheet through the Add Picture Method.')

		# close without saving the workbook

		$wb.SaveAs($loc + "Excel\Host\" + $server + "_VMWareStats" + $currentDate + ".xls")
		$wb.Save()
		$xl.Quit()
	}else {
		echo("	Files do not exist in Image directory please run graphVmWareStats first")
	}
}