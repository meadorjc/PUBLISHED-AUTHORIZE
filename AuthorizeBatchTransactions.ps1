# ------------------------------------------------------------------------------------------
# Authorize.NET Transaction Batch Downloader
# MEADORJC AT GMAIL.COM 03/20/2015
#
# adapted from:
# https://etechgoodness.wordpress.com/2014/02/25/sort-a-windows-forms-listview-in-powershell-without-a-custom-comparer/
# https://technet.microsoft.com/en-us/library/ff730949.aspx
# ------------------------------------------------------------------------------------------

# ------------------------------------------------------------------------------------------
# ASSEMBLIES
# ------------------------------------------------------------------------------------------
Add-Type -AssemblyName System.Windows.Forms

# ------------------------------------------------------------------------------------------
# CONSTANTS
# ------------------------------------------------------------------------------------------
Set-Variable Uri 							-option Constant	-value 'https://api.authorize.net/xml/v1/request.api'
Set-Variable NameSpace 						-option Constant	-value @{dns="AnetApi/xml/v1/schema/AnetApiSchema.xsd"}
Set-Variable TransactionFilterFile			-option Constant	-value ".\TransactionDetailsFilter.config"
Set-Variable MerchantAuthenticationFile 	-option Constant	-value ".\MerchantAuthentication.config"
Set-Variable SettledBatchListRequestFile 	-option Constant	-value ".\getSettledBatchListRequest.xml"
Set-Variable TransactionListRequestFile 	-option Constant	-value ".\getTransactionListRequest.xml"
Set-Variable TransactionDetailsRequestFile	-option Constant	-value ".\getTransactionDetailsRequest.xml"

# ------------------------------------------------------------------------------------------
# FUNCTION DEFINITIONS
# ------------------------------------------------------------------------------------------

# ------------------------------------------------------------------------------------------
# Get the xml request contents from file
# ------------------------------------------------------------------------------------------
function Get-RequestBody([string]$TransactionType)
{
	switch($TransactionType) {
		"SettledBatchListRequest" 	{ $Body = Get-Content $SettledBatchListRequestFile 	| Out-String; break}
		"TransactionListRequest" 	{ $Body = Get-Content $TransactionListRequestFile 	| Out-String; break}
		"TransactionDetailsRequest"	{ $Body = Get-Content $TransactionDetailsRequestFile | Out-String; break}
	}
	return $Body
}
# ------------------------------------------------------------------------------------------
# Merchant Auth codes stored in file
# ------------------------------------------------------------------------------------------
function Get-MerchantAuthentication([string]$File)
{
	$File = Get-Content $File | Out-String
	$Name  = Get-FilteredValue $File "name"
	$TransactionKey = Get-FilteredValue $File "transactionKey"
	
	return New-object PSObject -Property @{
		Name = $Name
		TransactionKey = $TransactionKey
	}
}
# ------------------------------------------------------------------------------------------
# POST and GET info from Authorize.NET
# ------------------------------------------------------------------------------------------
function Get-AuthorizeXmlResponseString([string]$Uri, [string]$Body)
{  
	# Convert/Encode
	$BodyBytes = [System.Text.Encoding]::UTF8.GetBytes($Body)

	# Create request, post, get, read
	$Request = [System.Net.WebRequest]::Create($Uri)
	$Request.Method="POST"
	# $Request.Proxy = $Null # Only needed if you use a proxy
	$Request.ContentType = 'application/xml'
	$RequestStream = $Request.GetRequestStream()
	$RequestStream.Write($BodyBytes, 0, $BodyBytes.Length)
	$RequestStream.Close()
	$Response = $Request.GetResponse()
	$ResponseStream = $Response.GetResponseStream()
	$ReadStream = New-Object System.IO.StreamReader $ResponseStream
	$Data = $ReadStream.ReadToEnd()
	$ResponseStream.Close()
	$ReadStream.Close()

	return $Data
}
# ------------------------------------------------------------------------------------------
# GUI for batch select
# ------------------------------------------------------------------------------------------
function Define-BatchSelectForm() {
	# Set up the environment
	$LastColumnClicked = 0 # tracks the last column number that was clicked
	$LastColumnAscending = $False # tracks the direction of the last sort of this column
 
	# Create a form, Configure the form
	$Form = New-Object System.Windows.Forms.Form
	$Form.Text = "Authorize.Net Batches"
	$Form.Size = New-Object System.Drawing.Size(500,600) 
	$Form.StartPosition = "CenterScreen"
	$Form.ControlBox = $False
	$Form.KeyPreview = $True

	# Create OK Button
	$OKButton = New-Object System.Windows.Forms.Button
	$OKButton.Size = New-Object System.Drawing.Size(75,25)
	$OKButton.Location = New-Object System.Drawing.Size($($($Form.Size.Width/2)-$($($OKButton.Size.Width))), $($Form.Size.Width-10))
	$OKButton.Anchor = "Bottom"
	$OKButton.Text = "OK"
	$OKButton.Add_Click({$BatchSelectForm.Close()})
	$Form.Controls.Add($OKButton)

	# Create Cancel Button
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Size = New-Object System.Drawing.Size(75,25)
	$CancelButton.Location = New-Object System.Drawing.Size($($($Form.Size.Width/2)), $($Form.Size.Width-10))
	$CancelButton.Text = "Cancel"
	$CancelButton.Anchor = "Bottom"
	$CancelButton.Add_Click({$BatchSelectForm.Close(); Stop-Process $Pid})
	$Form.Controls.Add($CancelButton) 

	# Define 'Enter' and 'Esc' keys
	$Form.Add_KeyDown({if ($_.KeyCode -eq "Enter") 
		{ $BatchSelectForm.Close()}})
	$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape") { $BatchSelectForm.Close(); Stop-Process $Pid}})
		
	# Configure the ListView
	$ListView = New-Object System.Windows.Forms.ListView
	$ListView.View = [System.Windows.Forms.View]::Details
	$ListView.Width = $Form.ClientRectangle.Width
	$ListView.Height = $($($OKButton.Location.Y)-5)
	$ListView.Anchor = "Top, Left, Right, Bottom"
	$ListView.Sorting = [System.Windows.Forms.SortOrder]::Descending
	$ListView.Name = "ListView"
	# Add the ListView to the Form
	$Form.Controls.Add($ListView)

	# Add columns to the ListView
	[void] $ListView.Columns.Add("Batch ID", 				-2, [System.Windows.Forms.HorizontalAlignment]::Center)
	[void] $ListView.Columns.Add("Settlement Time", -2, [System.Windows.Forms.HorizontalAlignment]::Center)
	[void] $ListView.Columns.Add("Charge Count", 		-2, [System.Windows.Forms.HorizontalAlignment]::Center)
	[void] $ListView.Columns.Add("Decline Count", 	-2, [System.Windows.Forms.HorizontalAlignment]::Center)
	[void] $ListView.Columns.Add("Total Count", 		-2, [System.Windows.Forms.HorizontalAlignment]::Center)
	[void] $ListView.Columns.Add("Total Amount", 		-2, [System.Windows.Forms.HorizontalAlignment]::Center)
	
	return $Form
}
# ------------------------------------------------------------------------------------------
# Add batch stats to listitems
# ------------------------------------------------------------------------------------------
function Add-ListViewItem($Form, $BatchTotals){
	# Add list items to ListView
	$ListViewItem = New-Object System.Windows.Forms.ListViewItem($BatchTotals.BatchId.ToString())
	[void] $ListViewItem.Subitems.Add([System.DateTime]::Parse($BatchTotals.SettlementTimeLocal).ToString("MM/dd/yy HH:mm:ss"))
	[void] $ListViewItem.Subitems.Add($BatchTotals.TotalChargeCount.ToString())
	[void] $ListViewItem.Subitems.Add($BatchTotals.TotalDeclineCount.ToString())
	[void] $ListViewItem.Subitems.Add($BatchTotals.TotalCount.ToString())
	[void] $ListViewItem.Subitems.Add($BatchTotals.TotalChargeAmount.ToString("C"))
	
	[void] $Form.Controls["ListView"].Items.Add($ListViewItem)
	[void] $Form.Controls["ListView"].AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::ColumnContent)
	[void] $Form.Controls["ListView"].AutoResizeColumns([System.Windows.Forms.ColumnHeaderAutoResizeStyle]::HeaderSize)

}
# ------------------------------------------------------------------------------------------
# GUI for download status
# ------------------------------------------------------------------------------------------
function Define-TransactionForm() {
	# Create a form, Configure the form
	$Form = New-Object System.Windows.Forms.Form
	$Form.Text = "Authorize.Net Batches"
	$Form.Size = New-Object System.Drawing.Size(200, 100) 
	$Form.StartPosition = "CenterScreen"
	$Form.ControlBox = $False
	$Form.KeyPreview = $True
	
	#Create Cancel Button
	$CancelButton = New-Object System.Windows.Forms.Button
	$CancelButton.Size = New-Object System.Drawing.Size(75,25)
	$CancelButton.Location = New-Object System.Drawing.Size($($($Form.Size.Width/2)-$($OKButton.Size.Width+10)), $($Form.Size.Width-10))
	$CancelButton.Text = "Cancel"
	$CancelButton.Add_Click({$TransactionForm.Close(); Stop-Process $Pid})
	$Form.Controls.Add($CancelButton) 
	
	$ObjLabel = New-Object System.Windows.Forms.Label
	$ObjLabel.Size = New-Object System.Drawing.Size(200,15) 
	$ObjLabel.Location = New-Object System.Drawing.Size(10, 10)
	$ObjLabel.Text = "Downloading Data"
	$ObjLabel.Name = "ObjLabel"
	$Form.Controls.Add($ObjLabel) 
	$Form.Add_KeyDown({if ($_.KeyCode -eq "Escape") {$TransactionForm.Close(); Stop-Process $Pid}}) 
	return $Form
}
# ------------------------------------------------------------------------------------------
# Get list of filters from file
# ------------------------------------------------------------------------------------------
function Get-TransactionFilters()
{
	$File = Get-Content $TransactionFilterFile
	
	$Filters = @()
	foreach($Line in $File) {
		$Line = $Line -replace '[\f\n\r\t\v]',''
		if ($Line[0] -ne '#' -and $Line.length -gt 0){
			$Filters += $Line
		}
	}
	
	return $Filters
}
# ------------------------------------------------------------------------------------------
# Get transaction data from XML
# ------------------------------------------------------------------------------------------
function Get-TransactionRecord([String]$Data, [System.Object]$Filters) {
    
	$Record = New-Object PSObject
	foreach($XPath in $Filters){
		$Record | add-member Noteproperty $XPath.Replace("/", "_")	$(Get-FilteredValue $Data $XPath)
	}
	return $Record
}
# ------------------------------------------------------------------------------------------
# Get transaction data from XML
# ------------------------------------------------------------------------------------------
function Get-FilteredValue([string]$Data, [string]$XPath){
	$DNSPath = "//dns:"+$XPath.Replace("/", "/dns:")
	$Value = $(Select-Xml -content $Data -xpath $DNSPath -Namespace $NameSpace).Node.InnerText
	if ($Value -eq $null){
		$Value = ""
	}
	
	return $Value
}

# ------------------------------------------------------------------------------------------
# Get batch info to be displayed
# ------------------------------------------------------------------------------------------
function Get-BatchTotals($Batch) {
	$TotalChargeCount = 0
	$TotalChargeAmount = 0
	$TotalDeclineCount = 0
	for($I = 0; $I -lt 3; $I++)
	{
		$TotalChargeCount += $Batch.Node.Statistics.Statistic[$I].ChargeCount
		$TotalChargeAmount += $Batch.Node.Statistics.Statistic[$I].ChargeAmount
		$TotalDeclineCount += $Batch.Node.Statistics.Statistic[$I].DeclineCount
	}
	
	return New-object PSObject -Property @{
		BatchId = $Batch.Node.BatchId
		SettlementTimeLocal = $Batch.Node.SettlementTimeLocal
		TotalChargeCount = $TotalChargeCount
		TotalDeclineCount = $TotalDeclineCount
		TotalCount = $TotalChargeCount + $TotalDeclineCount
		TotalChargeAmount = $TotalChargeAmount
	}
}
# ------------------------------------------------------------------------------------------
# Format records:
# ------------------------------------------------------------------------------------------
function Format-TransactionRecord($TransactionRecord){
	$RecordCols = ""
  
	$RecordProperties = $TransactionRecord | GM -MemberType Properties | Select -ExpandProperty Name
	[array]::Reverse($RecordProperties)
	foreach ($Property in $RecordProperties) {
		$RecordCols += $TransactionRecord.$Property + ","
	}
	
	return $RecordCols
}

#----------------------------------------------------------------------------------------------
# MAIN SCRIPT
#----------------------------------------------------------------------------------------------

# Define form and get reference to the form object
$BatchSelectForm = Define-BatchSelectForm
# Authentication codes
$MerchantAuthentication = Get-MerchantAuthentication $MerchantAuthenticationFile

# Get all batches within the last month
$Today = Get-Date; 
$LastMonth = $Today.AddMonths(-1);

# Set parameters to send to Authorize.NET
[xml]$Body = Get-RequestBody "SettledBatchListRequest"
$Body.getSettledBatchListRequest.merchantAuthentication.name = $MerchantAuthentication.Name.ToString()
$Body.getSettledBatchListRequest.merchantAuthentication.transactionKey = $MerchantAuthentication.TransactionKey.ToString()
$Body.getSettledBatchListRequest.includeStatistics = "true"
$Body.getSettledBatchListRequest.firstSettlementDate = $($Lastmonth.ToString("yyyy-MM-dd'T'HH:mm:ss"))
$Body.getSettledBatchListRequest.lastSettlementDate = $($Today.ToString("yyyy-MM-dd'T'HH:mm:ss"))

# POST and GET to ANET
[string]$Data = Get-AuthorizeXmlResponseString $Uri $Body.OuterXml

# Get multiple batch nodes
$Batches = Select-Xml -content $Data -xpath "//dns:batch" -namespace $NameSpace

# Calculate batch info and add to list
foreach($Batch in $Batches){
	$BatchTotals = Get-BatchTotals $Batch
	Add-ListViewItem $BatchSelectForm $BatchTotals
}

# Show the form - Script will stop at ShowDialog() until user selects a batch
[void] $BatchSelectForm.ShowDialog()

#Define 2nd form and get reference to the form object
$TransactionForm = Define-TransactionForm
$TransactionForm.Show()

# Get the lists of transactions for selected batches
[xml]$Body = Get-RequestBody "TransactionListRequest"
$Body.getTransactionListRequest.merchantAuthentication.name = $MerchantAuthentication.Name.ToString()
$Body.getTransactionListRequest.merchantAuthentication.transactionKey = $MerchantAuthentication.TransactionKey.ToString()
$Transactions = @()
$TotalRecordCount = 0;
foreach ($SelectedItem in $BatchSelectForm.Controls["Listview"].SelectedItems) {

	# Set parameters and send to Authorize.NET
	$Body.getTransactionListRequest.batchId = $SelectedItem.Text
	
	# Post and Get to ANET
	[string]$Data = Get-AuthorizeXmlResponseString $Uri $Body.OuterXml
	
	# Add the "Charge Count"(2) and "Decline Count"(3) together for each batch from the listItems
	$TotalRecordCount += [int]::Parse($SelectedItem.SubItems[2].Text) + [int]::Parse($SelectedItem.SubItems[3].Text)
	$Transactions += select-xml -content $Data -xpath "//dns:transaction" -namespace $NameSpace | select-object -exp node
}

#Filters for transaction XML data
$Filters = Get-TransactionFilters

# Get the details of each transaction
[xml]$Body = Get-RequestBody "TransactionDetailsRequest"
$Body.getTransactionDetailsRequest.merchantAuthentication.name = $MerchantAuthentication.Name.ToString()
$Body.getTransactionDetailsRequest.merchantAuthentication.transactionKey = $MerchantAuthentication.TransactionKey.ToString()
$TransactionRecordList = @()

#Create Header	
$TransactionHeaders = New-Object PSObject
$Filters | %{ $($TransactionHeaders | add-member Noteproperty $_.Replace("/", "_")	$_.Replace("/","_")) }
$TransactionRecordList += $TransactionHeaders

# Get transaction details from ANET	
$Count = 0
foreach ($Transaction in $Transactions){
	$Count++
	# Set parameters and send to Authorize.NET
	$Body.getTransactionDetailsRequest.transId = $Transaction.TransId.ToString()
	
	# Post and Get to ANET
	[string]$Data = Get-AuthorizeXmlResponseString $Uri $Body.OuterXml
	$TransactionRecordList += Get-TransactionRecord $Data $Filters

	# Update the form counter
	$TransactionForm.Controls["ObjLabel"].Text = "Downloading: $($Count) / $($TotalRecordCount)"
	$TransactionForm.Update()
}

# Aggregate data in list of formatted tuples
$OutTupleList = @(); $Count = 0;
foreach ($TransactionRecord in $TransactionRecordList){
	$Count++
	$OutTupleList += Format-TransactionRecord $TransactionRecord
	$TransactionForm.Controls["ObjLabel"].Text = "Processing: $($Count) / $($TotalRecordCount)"
	$TransactionForm.Update()
}

# Define FILENAME and write to file
$FileName = "AuthorizeTransactions_$($(get-date -format 'MM-dd-yyyy'))"
Out-File -filepath ".\$($FileName).csv" -inputObject $OutTupleList -encoding ASCII
$TransactionForm.Close()

# Close the shell; shell can hang otherwise
Stop-Process $Pid