##############
## COPYSPAJ ##
##############

function copySPAJ 
{
    Param ([String]$depart, [string]$area, [String]$loc)

  # date
$date = (Get-Date).ToString("yyyy-MM") 
$day = (Get-Date).ToString("dd")
$reportdate = (Get-Date).AddDays(-1).ToString("yyyyMMdd")

  # lokasi spaj di lokal
$path = "\\eli-fs-node01\spaj$"
#$path = "D:\WORK\Task - project\Underwriting-SPAJ\SFTP"
$childpath = "$date\12\$depart\$area\$loc"
$localpath = "$path\01_SPAJ_SOA\$childpath"

  # path generated report
$pathreport = "$path\01_SPAJ_SOA\$date\12"

# SFTP SESSION
   # authentication
$passwd = ConvertTo-SecureString "Policy123!" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ("sys-ftp-policy-printing", $passwd)

    # Path folder spaj di SFTP
$pathspajsoa1 = "/_SPAJ_ESUBMISSION/test/01_SPAJ_SOA/$date/12/$depart/$area/$loc"
$pathspajpending1 = "/_SPAJ_ESUBMISSION/test/02_SPAJ_PENDING_NBA/$date/12/Agency/"
$pathspajpending2 = "/_SPAJ_ESUBMISSION/test/02_SPAJ_PENDING_NBA/$date/12/Inbranch/"
$pathspajpending = @($pathspajpending1,$pathspajpending2)

    # create session SFTP 
$sftpsession = New-SFTPSession -ComputerName sftp-uat.myequity.id -Credential $creds -AcceptKey

 # COPY SPAJ DARI FOLDER 01_SPAJ_SOA  
 $folderspaj = Get-ChildItem "$localpath" | Select-Object FullName 

   # cek apakah folder 01_SPAJ_SOA nya sudah ada di sftp atau belum
   if (!(Test-SFTPPath -Session $sftpsession -Path $pathspajsoa1)) { 
     New-SFTPItem -Session $sftpsession -Path $pathspajsoa1 -ItemType Directory -Recurse
   }
   # copy spaj ke folder 01_SPAJ_SOA di sftp
   foreach ($a in $folderspaj) { 
     Set-SFTPItem -SessionId $sftpsession.SessionID -Path $a.FullName -Destination $pathspajsoa1
   }

######################################################################################################################################

 # CREATE FOLDER DI FOLDER 02_SPAJ_PENDING_NBA DI SFTP  & COPY FILE DI FOLDER 02_SPAJ_PENDING_NBA DI SFTP KE LOKAL 

     # cek apakah folder sudah terbentuk atau blm
    $pathspajpending | ForEach-Object { 
      if ($_ -and !(Test-SFTPPath -Session $sftpsession -Path $_)) { 
        # buat folder
        New-SFTPItem -Session $sftpsession -Path $_ -ItemType Directory -Recurse
      }
    }

   # copy file dengan case setelahnya hanya copy file yang belum pernah tercopy sebelumnya/baru ke lokal
    # setelah me-copy file/folder ke lokal, simpan nama file nya kedalam file/array 
    # lalu setelahnya sebelum copy cek nama file nya sudah pernah di copy atau blm, jika belum copy jika sudah jgn copy 

      $recordfile = "/_SPAJ_ESUBMISSION/test/record.txt"

   # Membaca file yang berisi nama file yang sudah disalin sebelumnya
     $record = @()
     if (Test-SFTPPath -SessionId $sftpsession.SessionID -Path $recordfile) {
       $record = Get-SFTPContent -SessionId $sftpsession.SessionID $recordfile
     } else { 
         New-SFTPItem -Session $sftpsession -Path $recordfile -ItemType Directory 
         $record = Get-SFTPContent -SessionId $sftpsession.SessionID $recordfile
       }

$filespajpending1 = Get-SFTPChildItem -Session $sftpsession -Path $pathspajpending1
$filespajpending2 = Get-SFTPChildItem -Session $sftpsession -Path $pathspajpending2

$localpathspajpending1 = "$path\02_SPAJ_PENDING_NBA\12\$day\Agency\"
$localpathspajpending2 = "$path\02_SPAJ_PENDING_NBA\12\$day\Inbranch\"

#COPY FILE DARI FOLDER 02_SPAJ_PENDING_NBA - AGENCY" 
 foreach ($c in $filespajpending1) { 
   if ($record -match $c.Name) {
     Write-Host "File $($c.Name) sudah ada di dalam record dan sudah ada di older Agency."
     } else {
         Get-SFTPItem -SessionId $sftpsession.SessionID -Path $c.FullName -Destination $localpathspajpending1 
         Write-Host "File $($c.Name) telah disalin ke folder Agency."
         $record += "$($c.Name)`r`n"
       }
    }

#COPY FILE DARI FOLDER 02_SPAJ_PENDING_NBA - INBRANCH"
 foreach ($d in $filespajpending2) { 
   if ($record -match $d.Name) {
     Write-Host "File $($d.Name) sudah ada di dalam record dan sudah ada di older Inbranch."
     } else {
         Get-SFTPItem -SessionId $sftpsession.SessionID -Path $d.FullName -Destination $localpathspajpending2 
         Write-Host "File $($d.Name) telah disalin ke folder Inbranch."
         $record += "$($d.Name)`r`n"
       }
    }

 # Menyimpan kembali nama file yang sudah disalin ke dalam file record
Set-SFTPContent -SessionId $sftpsession.SessionID -Path $recordfile -Value ($record -join [Environment]::NewLine) 

#putus session sftp
Remove-SFTPSession -Session $sftpsession

######################################################################################################################################

 Start-Sleep -Seconds 3

# GENERATED FILE CSV REPORT
    #$csv1 = Join-Path -Path $path -Childpath "01_SPAJ_SOA\$date\$day\report_01_SPAJ_SOA_$($reportdate).csv"
    $csv1 = Join-Path -Path $path -Childpath "01_SPAJ_SOA\$date\12\report_01_SPAJ_SOA_$($reportdate).csv"
    
    if (!(Test-Path -Path $csv1)) {
        New-Item -Path $csv1 -ItemType File
       }
    
    $reportdate = (Get-Date).AddDays(-1).ToString("yyyyMMdd")
    $nama =  @{N='Nama SPAJ'; Expression={$_.Name}}
    $tanggal = @{Name='Tanggal';Expression={"$reportdate"}}
    $department = @{Name='Department';Expression={"$depart"}}
    $areas = @{Name='Area';Expression={"$area"}}
    $cabang = @{Name='Cabang';Expression={"$loc"}}

    $fileGen1 = Get-ChildItem "$localpath" -Exclude *.csv* | Select-Object $nama, $tanggal, $department, $areas, $cabang | Export-Csv "$csv1" -Delimiter ',' -NoTypeInformation -Append
}

copySPAJ -depart "Agency" -area 'Area I' -loc "BN Semarang"
copySPAJ -depart "Agency" -area 'Area I' -loc "Indramayu"
copySPAJ -depart "Agency" -area 'Area II' -loc "Denpasar"
copySPAJ -depart "Agency" -area 'Area II' -loc "Makassar"
copySPAJ -depart "Agency" -area 'Area II' -loc "Medan"
copySPAJ -depart "Agency" -area 'Area II' -loc "Palembang"
copySPAJ -depart "Agency" -area 'Area III' -loc "Bandung"
copySPAJ -depart "Agency" -area 'Area III' -loc "Cirebon"
copySPAJ -depart "Agency" -area 'Area III' -loc "Head Office"
copySPAJ -depart "Agency" -area 'Area III' -loc "Jakarta"
copySPAJ -depart "Agency" -area 'Area IV' -loc "Kediri"
copySPAJ -depart "Agency" -area 'Area IV' -loc "Magelang-Yogja"
copySPAJ -depart "Agency" -area 'Area IV' -loc "Malang"
copySPAJ -depart "Agency" -area 'Area IV' -loc "Solo"
copySPAJ -depart "Agency" -area 'Area IV' -loc "Surabaya"
copySPAJ -depart "Agency" -area 'Area IV' -loc "Timika"
copySPAJ -depart "Inbranch" -area 'Area I' -loc "Bali"
copySPAJ -depart "Inbranch" -area 'Area II' -loc "Jakarta"
copySPAJ -depart "Inbranch" -area 'Area III' -loc "Bandung"
copySPAJ -depart "Inbranch" -area 'Area IV' -loc "Semarang"
copySPAJ -depart "Inbranch" -area 'Area IV' -loc "Solo"
copySPAJ -depart "Inbranch" -area 'Area IV' -loc "Surabaya"

