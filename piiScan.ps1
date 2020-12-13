function findInMSWord(){
  $Word = New-Object -ComObject Word.Application
  $Document = $Word.Documents.Open("C:\Users\exaybachay\Programming\FindPII\PII.docx")
  $Document.Paragraphs | ForEach-Object {
    $_.Range.Text
  }

  #clean up stuff
  #[System.Runtime.InteropServices.Marshal]::ReleaseComObject($range) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($document) | Out-Null
  [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
  Remove-Variable -Name word
  Remove-Variable -Name Document
  
  [gc]::collect()
  [gc]::WaitForPendingFinalizers()
}

function findInMSExcel(){
  $path = (pwd).path
  $excelSheets = Get-Childitem -Path $path -Include *.xls,*.xlsx -Recurse

  $file = "$pwd\pii.xlsx"
  $SearchString = "March 5, 1999"

  $Excel = New-Object -ComObject Excel.Application
  #$Excel.visible = $false
  $Workbook = $Excel.Workbooks.Open($file)

  $Worksheets = $Workbooks.worksheets
  $Worksheet = $workbook.Worksheets.Item(1)

  $Range = $Worksheet.Range("A:Z")
  $Search = $Range.find($SearchString)

  $excel.quit | out-null
  #write-output $search
}

#create some patterns to look for
$ssn = '\b[0-9]{3}\-[0-9]{2}\-[0-9]{4}\b'
$bday = '\b[0-9]{1,2}(-|\/)[0-9]{1,2}(-|\/)[0-9]{4}\b'
$eurobday = '\b[0-9]{4}(-|\/)[0-9]{1,2}(-|\/)[0-9]{1,2}\b'
$fullbday = '\b(January|February|March|April|May|June|July|August|September|October|November|December)\s[0-9]{1,2}.\s[0-9][0-9][0-9][0-9]\b'
$eurofullbday = '\b[0-9]{1,2}\s(January|February|March|April|May|June|July|August|September|October|November|December).\s[0-9][0-9][0-9][0-9]\b'
$allmatches = '(\b[0-9]{4}(-|\/)[0-9]{1,2}(-|\/)[0-9]{1,2}\b)|(\b[0-9]{1,2}(-|\/)[0-9]{1,2}(-|\/)[0-9]{4}\b)|(\b[0-9]{3}\-[0-9]{2}\-[0-9]{4}\b)|(\b(January|February|March|April|May|June|July|August|September|October|November|December)\s[0-9]{1,2}.\s[0-9][0-9][0-9][0-9]\b)|(\b[0-9]{1,2}\s(January|February|March|April|May|June|July|August|September|October|November|December).\s[0-9][0-9][0-9][0-9]\b)'

#look for matches in current directory and subdirectories
#Get-ChildItem -Recurse | Get-Content | where { $_ | Select-String -Pattern $allmatches }

#look for matches in current directory and subdirectories
$files = Get-ChildItem -Recurse
foreach ($file in $files){
  $foundinfo = Get-Content $file | where {
    $_ | Select-String -Pattern $allmatches
  }
  if($foundinfo){
    write-output $file.name
    write-output $foundinfo
    write-output " "
  }
}
