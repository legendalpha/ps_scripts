


#$user = ""
#$KEYS_DIR = "..\_KEYS_"

#$username = "$user"
#$password = Get-Content $KEYS_DIR\$user.pass 
#$key = Get-Content $KEYS_DIR\AES.key 
#$cred = New-Object -TypeName System.Management.Automation.PSCredential -argumentlist $username, ($password | ConvertTo-SecureString -Key $key)
#$ErrorActionPreference = 'SilentlyContinue'

#$Computer = Read-Host Please Enter Host name
#  Get-WMIObject Win32_NetworkAdapterConfiguration -Computername $Computer -credential $cred| `
#  Where-Object {$_.IPEnabled -match "True"} | `
#  Select-Object -property DNSHostName,ServiceName,@{N="DNSServerSearchOrder";
#              E={"$($_.DNSServerSearchOrder)"}},
#              @{N='IPAddress';E={$_.IPAddress}},
#              @{N='DefaultIPGateway';E={$_.DefaultIPGateway}} | FT

function Get_IP($server) {
  logger "Initiating the connection to $server"

  $output = Invoke-Command -computername $server -credential $cred -ScriptBlock { ipconfig /all }

  if ( $? -eq $False ) {
    logger("Unable to connect or retrieve network information for $server")
  }
  else {
    logger("Output obtained for $server. See $server.log for more details")
    
    Write-Output $output  > "$server.log"
  }
}

function logger($msg) {
  $dt = Get-Date
  $masteroutput = "Output.log"
  Write-Output "$dt : $msg" >> $masteroutput
}

function Get-FileName($initialDirectory) {  
  [System.Reflection.Assembly]::LoadWithPartialName("system.windows.forms") |
  Out-Null

  $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
  $OpenFileDialog.initialDirectory = $initialDirectory
  $OpenFileDialog.filter = "Excel Workbook (*.xlsx)| *.xlsx"
  $OpenFileDialog.ShowDialog() | Out-Null
  $OpenFileDialog.filename
}

function Get-ListFrom-Excel() {
  $objExcel = New-Object -ComObject Excel.Application
  $objExcel.Visible = $false
  $WorkBook = $objExcel.Workbooks.Open($INPUT_FILE_NAME)
  #$WorkBook.sheets | Select-Object -Property Name
  $count = $workbook.sheets.Count

  $list = @()
  $column = 1
  $start_row = 1
  for ($i = 1 ; $i -le $count ; $i++  ) {
    $SHEET_NAME = $WorkBook.Sheets.Item($i).Name
    if ( $SHEET_NAME -eq "Sheet1" ) {
      $WorkSheet = $WorkBook.Sheets.Item($i)
      $totalNoOfRecords = ($WorkSheet.UsedRange.Rows).count
      #$totalNoOfRecords
      for ($row = $start_row; $row -le $totalNoOfRecords ; $row ++) {
        $value = $WorkSheet.Cells($row, $column).value2
        $list += $value
      }
      #$list = $WorkSheet.Range("A1").copy()
    }
    else {
      #Write-Output "Spreadsheet name is not 'Sheet1'. Exiting..."
      #exit 99
      Continue
    }
  }

  ## Clean UP

  $objExcel.quit()
  [System.Runtime.Interopservices.Marshal]::ReleaseComObject($objExcel) | Out-Null
  $objExcel = $null
  return $list
}

$INPUT_FILE_NAME = Get-FileName -initialDirectory $env:USERPROFILE
$SERVER_LIST = Get-ListFrom-Excel($INPUT_FILE_NAME)

#$hostnames = @("10.100.2.80")

foreach ( $server in $SERVER_LIST ) {
  "this is $server"
  #Get_IP "$server"
}