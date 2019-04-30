##########################
#   CLI-Emailer Script   #
##########################

# 1. Choose your Excel list of students
# 2. Choose the template you wish to use
# 3. Choose to send test email to your own email account
# 4. Send template emails to all students on list


#############
# Functions #
#############

# shows dialog box to choose file
Function Get-FileName($initialDirectory) {
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $OpenFileDialog.InitialDirectory = $initialDirectory
    $OpenFileDialog.ShowDialog() | Out-Null
    $OpenFileDialog.filename
}


# returns hashtable of data from Excel file
Function Get-EmailList($fileDirectory) {
    $objExcel = New-Object -ComObject Excel.Application
    $objExcel.Visible = $false
    $workbook = $objExcel.Workbooks.Open($fileDirectory)
    $sheetName = "Sheet1"
    $sheet = $workbook.Worksheets.Item($sheetName)

    $rowMax = ($sheet.UsedRange.Rows).count

    $rowID, $colID = 2, 1
    $rowTo, $colTo = 2, 2
    $rowFirstName, $colFirstName = 2, 3
    $rowLastName, $colLastName = 2, 4

    $emailList = [ordered]@{}

    # creates hashtables for each row in Excel file. 

    #     $emailList = {
    #        "id1": {to:x; from:x; ...}
    #        "id2": {to:x; from:x; ...}
    #    }

    for ($i = 0; $i -le $rowMax - 2; $i++) {
    
        $emailHash = [ordered]@{}

        $id = $sheet.Cells.Item($rowID + $i, $colID).text
        $to = $sheet.Cells.Item($rowTo + $i, $colTo).text
        $firstName = $sheet.Cells.Item($rowFirstName + $i, $colFirstName).text
        $lastName = $sheet.Cells.Item($rowLastName + $i, $colLastName).text

        $emailHash.add("to", $to)
        $emailHash.add("firstname", $firstName)
        $emailHash.add("lastname", $lastName)

        $emailList.add($id, $emailHash)

    }

    $workbook.Close()
    $objExcel.Quit()

    return $emailList
}



#############
# Variables #
#############


# $directory = Get-Location
$sheetName = "Sheet1"
$excelFileName = Get-FileName       # gets filename of excel list
$emailAccount = "name@email.edu"
$signaturePath = "Signature.htm"
$HTMLsignature = Get-Content "$env:USERPROFILE\AppData\Roaming\Microsoft\Signatures\$signaturePath"
$templateFileName = Get-FileName

$listContents = Get-EmailList($excelFileName) # gets contents of excel file


#############
#   Logic   #
#############


# 1. Ask to send test email, then send if Y
$testResponse = Read-Host "Send test email to" $emailAccount "? (y/n)"

If ($testResponse -eq "y") {
    
    "Sending test email to $emailAccount... "

    $Outlook = New-Object -ComObject Outlook.Application
    $Mail = $Outlook.CreateItemFromTemplate("$templateFileName")
    $Mail.To = $emailAccount
    $Mail.Sender = $emailAccount
    $Mail.HTMLBody = $Mail.HTMLBody.Replace("{name}", "TEST")
    # for some reason, concatenating the signature like this messes up the formatting
    # $Mail.HTMLBody = $Mail.HTMLBody + $HTMLsignature

    try {
        $Acct = $Outlook.Session.Accounts.Item($emailAccount)
        $Mail.SendUsingAccount = $acct
        
        $Mail.Send()
        Write-Host "Sent to " $emailAccount " from " $emailAccount
    }
    catch {
        Write-Host "Error"
    }
}
Else {
    "Test email not sent."
}

# 2. Ask to send emails to X number of emails as show in Excel file, send if Y
Write-Host " "
Write-Host "Emails queued to be sent:"
Write-Host " "
foreach ($val in $listContents.values) {
    "To: " + $val.firstname + " (" + $val.to + ") - From: " + $emailAccount
}


$c = $listContents.Count
Write-Host " "
$response = Read-Host "Send $c emails? (y/n)"



If ($response -eq "y") {
    "Sending to..."
    # Create Email from Template
    $Outlook = New-Object -ComObject Outlook.Application

    foreach ($val in $listContents.values) {
        $val.firstname + " (" + $val.to + ")"
        $Mail = $Outlook.CreateItemFromTemplate("$templateFileName")
        $Mail.Sender = $emailAccount
        $Mail.To = $val.to
        $Mail.HTMLBody = $Mail.HTMLBody.Replace("{name}", $val.firstname)
        # $Mail.HTMLBody = $Mail.HTMLBody + $HTMLsignature
        $Mail.Send()
    }
}
Else {
    "Cancelled."
}
