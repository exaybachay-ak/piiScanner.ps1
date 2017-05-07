#create some patterns to look for
$ssn = '\b[0-9]{3}\-[0-9]{2}\-[0-9]{4}\b'
$bday = '\b[0-9]{1,2}(-|\/)[0-9]{1,2}(-|\/)[0-9]{4}\b'
$eurobday = '\b[0-9]{4}(-|\/)[0-9]{1,2}(-|\/)[0-9]{1,2}\b'
$fullbday = '\b(January|February|March|April|May|June|July|August|September|October|November|December)\s[0-9]{1,2}.\s[0-9][0-9][0-9][0-9]\b'
$eurofullbday = '\b[0-9]{1,2}\s(January|February|March|April|May|June|July|August|September|October|November|December).\s[0-9][0-9][0-9][0-9]\b'
$allmatches = '(\b[0-9]{4}(-|\/)[0-9]{1,2}(-|\/)[0-9]{1,2}\b)|(\b[0-9]{1,2}(-|\/)[0-9]{1,2}(-|\/)[0-9]{4}\b)|(\b[0-9]{3}\-[0-9]{2}\-[0-9]{4}\b)|(\b(January|February|March|April|May|June|July|August|September|October|November|December)\s[0-9]{1,2}.\s[0-9][0-9][0-9][0-9]\b)|(\b[0-9]{1,2}\s(January|February|March|April|May|June|July|August|September|October|November|December).\s[0-9][0-9][0-9][0-9]\b)'

#look for matches in current directory and subdirectories
Get-ChildItem -Recurse | Get-Content | where { $_ | Select-String -Pattern $allmatches }
