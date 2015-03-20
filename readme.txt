READ ME - meadorjc at gmail.com

Files Required:
	readme.txt													- you are here.
	AuthorizeBatchTransactions.ps1			- source code
	TransactionDetailsFilter.config			- specifies columns to include in csv file
	MerchantAuthentication.config				-	config file for Merchant Authentication information
	getSettledBatchListRequest.xml			- Authorize.NET API XML Request Schema
	getTransactionListRequest.xml       - Authorize.NET API XML Request Schema
	getTransactionDetailsRequest.xml    - Authorize.NET API XML Request Schema
	

Instructions:
1. 	Download AuthorizeBatchTransactions.zip and unzip all files into a folder.
2.	Insert organization's Merchant Authentication data in MerchantAuthentication.config
			Name (aka Code)
			Transaction Key
3. 	Open TransactionDetailsFilter.config and configure needed transaction details
4. 	Run the script*

*Script Execution Options

	ADD SIGNATURE: If you choose to add a signature to the script, you can set the global 
	execution policy to RemoteSigned or AllSigned through the following commands. Adding a 
	signature is recommended if you plan to deploy this script to multiple computers and/or users.
		
		Powershell.exe Set-ExecutionPolicy AllSigned
		-or-
		Powershell.exe Set-ExecutionPolicy RemoteSigned
		
	RUN UNRESTRICTED SESSION: If you choose to run the script as an unsigned, you can specify
	a single, unrestricted shell instance by running the following command.
		
		PowerShell.exe -ExecutionPolicy Unrestricted "PATH_TO_FILE\AuthorizeBatchTransactions.ps1"