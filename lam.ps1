
CLS

$intro =
"
  ______              _  _____                                
 |  ____|            | |/ ____|                               
 | |__  __  _____ ___| | (___   ___ _ __ __ _ _ __   ___ _ __ 
 |  __| \ \/ / __/ _ \ |\___ \ / __| '__/ _`  | '_ \ / _ \ '__|
 | |____ >  < (_|  __/ |____) | (__| | | (_| | |_) |  __/ |   
 |______/_/\_\___\___|_|_____/ \___|_|  \__,_| .__/ \___|_|   
                                             | |              
                                             |_|              
`tBETA VERSION 0.1

"
cls
write-host $intro -ForegroundColor Magenta
Do {
Write-Host "`tInsert link here" -nonewline -f Green; 
    $url = Read-Host " "
} While ($url -notmatch "boliga.dk/")

write-host "`t`tThanks! Preparing system:" -f green
Start-Sleep -Seconds 1
#Preparing modules
write-host "`t`t`t- Checking system requirements..." -f green
sleep -s 2  
write-host "`t`t`t- Checking modules..." -f green  
    if (!(Get-Module -ListAvailable -Name ImportExcel)) 
    {write-host "`t`t`tModule not found! Installing..." -f green;
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; sleep -s 1;
    Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | out-null; sleep -s 1;
    Install-Module -Name ImportExcel -Force;}

#Excel sheet
write-host "`t`t`t- Creating directory for output..." -f green
    new-item -Path "c:\ExcelScraper" -ItemType Directory -ea SilentlyContinue | Out-Null
    $date = get-date -f "yyyy-MM-dd-HH.mm.ss"
    sleep -s 1
    $file = "C:\ExcelScraper\$date.xlsm"
write-host "`t`t`t- Downloading template..." -f green
    Invoke-WebRequest -Uri "https://github.com/Andreas6920/project-lam/raw/main/Eksempel.xlsm" -OutFile $file -UseBasicParsing
    sleep -s 1

write-host "`t`tInitializing Program:" -f green
sleep -s 2

sleep -s 1 
write-host "`t`t`t- Pulling data..(this may take a while)" -f green
    #$link = "https://www.boliga.dk/salg/resultater?propertyType=3&salesDateMin=2018&zipcodeFrom=2610&zipcodeTo=2610&page=1&searchTab=1&sort=date-d&pageSize=1000"
    $link = $url
    $scrape = (Invoke-WebRequest -uri $link).Allelements
    $antal = ($scrape | where class -match "table-row white-background|table-row gray-background").Count -1

write-host "`t`t`t- Sorting data..." -f green
    #$adresse = (($scrape | where data-gtm -eq "sales_address").innerHTML| Foreach-object {$_ -replace '\<.*',""}).Trim()
    #$by = (($scrape | where data-gtm -eq "sales_address").innerHTML| Foreach-object {$_ -replace '.*\"">',""}).Trim()
    $fulladdress = (($scrape | where data-gtm -eq "sales_address").innerHTML | Foreach-object {$_ -replace '\<.*>',","})
    $Købesum = ($scrape | where class -eq "text-nowrap" | where innertext -match "kr.").innerText
    $Salgsdato = (($scrape | where class -eq "text-nowrap" | where innerHTML -match "\d{2}-\d{2}-\d{4}").innerText | Foreach-object {$_ -replace '-','/'}).Trim()
    $Boligtype = (($scrape | where class -eq "property-3 hide-text").innerText | Foreach-object {$_ -replace 'EEjerlejlighed',''}).Trim()
    $KRM2 = (($scrape | where class -eq "text-nowrap mt-1" | Where innerText -Match "kr\/m").innerText| Foreach-object {$_ -replace 'kr\/m²',''}).Trim()
    $Værelser = ($scrape | where class -eq "table-col text-center" | where outerText -match "(?<!\S)\d(?!\S)").InnerText
    $M2 = (($scrape | where class -eq "text-nowrap" | where innerText -match "m²").innerText | Foreach-object {$_ -replace 'm²',''}).Trim()
    $Byggeår = ($scrape | where class -eq "table-col text-center" | where innertext -match "^16\d{2}|^17\d{2}|^18\d{2}|^19\d{2}|^20\d{2}").innerText
    # % - UNDLAD TIL AT STARTE MED
    #AKTUEL VÆRDI
    
    

write-host "`t`t`t- Preparing data for Excel..." -f green
    $oversigt = @();
    0..$antal | % {$oversigt += New-Object -TypeName psobject -Property @{`
    Adresse=$fulladdress[$_].Trim();`
    Købesum=$Købesum[$_];`
    Salgsdato=$Salgsdato[$_];`
    Boligtype=$Boligtype[$_];`
    KRM2=$KRM2[$_];`
    Værelser=$Værelser[$_];`
    M2=$M2[$_];`
    Byggeår=$Byggeår[$_];`
 
    }}

$oversigt | select adresse,Købesum,Salgsdato,Boligtype,KRM2,Værelser,M2,Byggeår | Export-Excel -path $file -StartRow 2 -NoHeader -WorksheetName Boliga -Show