param($statsLoc)
$allvms = @()
$allhosts = @()
$currentDate = (get-date).ToString("yyyyMMdd")
$loc = $statsLoc
$csvLoc = $statsLoc + "Hosts.csv"

[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
$data = import-csv $csvLoc
$servers = $data | select -ExpandProperty HostName -Unique | sort

foreach ($server in $servers){
	$day = 1
	$chart = new-object System.Windows.Forms.DataVisualization.Charting.Chart
	$chartarea = new-object system.windows.forms.datavisualization.charting.chartarea
	$chartarea.AxisY.Minimum = 0
	$chartarea.AxisY.Maximum = 100
	$chart.width = 1500
	$chart.Height = 600
	$chart.Left = 20
	#$chartarea.BackColor = DCDCDC
	$chart.top = 30
	$chart.Name = $server
	$chart.ChartAreas.Add($chartarea)
	$legend = New-Object system.Windows.Forms.DataVisualization.Charting.Legend
	$chart.Legends.Add($legend)

	$points = $data | Where-Object {($_.HostName -eq $server)}
	$chart.Series.Add("Memory Max (%)")
	$chart.Series["Memory Max (%)"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
	$chart.Series.Add("Memory Min (%)")
	$chart.Series["Memory Min (%)"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
	$chart.Series.Add("CPU Max (%)")
	$chart.Series["CPU Max (%)"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
	$chart.Series.Add("CPU Min (%)")
	$chart.Series["CPU Min (%)"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Line
	
	foreach ($point in $points){
		$datapoint = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
		$datapoint.SetValueY($point.MemMax)
		$datapoint.AxisLabel = $point.Date
		$chart.Series["Memory Max (%)"].Points.Add($datapoint)
		
		$datapoint2 = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
		$datapoint2.SetValueY($point.MemMin)
		$datapoint2.AxisLabel = $point.Date
		$chart.Series["Memory Min (%)"].Points.Add($datapoint2)
		
		$datapoint3 = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
		$datapoint3.SetValueY($point.CPUMax)
		$datapoint3.AxisLabel = $point.Date
		$chart.Series["CPU Max (%)"].Points.Add($datapoint3)
		
		$datapoint4 = New-Object System.Windows.Forms.DataVisualization.Charting.DataPoint
		$datapoint4.SetValueY($point.CPUMin)
		$datapoint4.AxisLabel = $point.Date
		$chart.Series["CPU Min (%)"].Points.Add($datapoint4)
		$day = $day + 1
	}
	$filename = $loc + "Images\Host\" + $server + "_" + $currentDate + ".png"
	$chart.SaveImage($filename, "PNG")
}