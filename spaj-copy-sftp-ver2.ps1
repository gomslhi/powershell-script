##################
#### COPYSPAJ ####
##    gomsil    ##
##################

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
$childpath = "$date\$day\$depart\$area\$loc"
$localpath = "$path\01_SPAJ_SOA\$childpath"

  # path generated report
$pathreport = "$path\01_SPAJ_SOA\$date\$day"

 # SFTP SESSION
   # authentication
$passwd = ConvertTo-SecureString "Policy123!" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ("sys-ftp-policy-printing", $passwd)
   # create session SFTP 
$sftpsession = New-SFTPSession -ComputerName sftp-uat.myequity.id -Credential $creds -AcceptKey

 # Path folder spaj di SFTP
    # path folder 01_SPAJ_SOA
$pathspajsoa1 = "/_SPAJ_ESUBMISSION/test/01_SPAJ_SOA/$date/$day/$depart/$area/$loc"
    # path folder 02_SPAJ_PENDING_NBA
$pathspajpending1 = "/_SPAJ_ESUBMISSION/test/02_SPAJ_PENDING_NBA/$date/$day/Agency/$area/$loc"
$pathspajpending2 = "/_SPAJ_ESUBMISSION/test/02_SPAJ_PENDING_NBA/$date/$day/Inbranch/$area/$loc"
$pathspajpending = @($pathspajpending1,$pathspajpending2)
    # path folder 03_SPAJ_PROCESS_UW
$pathspajprocessuw1 = "/_SPAJ_ESUBMISSION/test/03_SPAJ_PROCESS_UW/$date/$day/Agency"
$pathspajprocessuw2 = "/_SPAJ_ESUBMISSION/test/03_SPAJ_PROCESS_UW/$date/$day/Inbranch"
$pathspajprocessuw = @($pathspajprocessuw1,$pathspajprocessuw2)

######################################################################################################################################

 # COPY SPAJ DARI FOLDER 01_SPAJ_SOA DI FTP KE SFTP  

 $folderspaj = Get-ChildItem "$localpath" | Select-Object FullName 

   # cek apakah folder 01_SPAJ_SOA nya sudah ada di sftp atau belum
   if (!(Test-SFTPPath -Session $sftpsession -Path $pathspajsoa1)) { 
     New-SFTPItem -Session $sftpsession -Path $pathspajsoa1 -ItemType Directory -Recurse
   }
   # tambahkan ke file record untuk file sudah dicopy ke sftp folder 1.soa
   $file_record_soa = "$path\01_SPAJ_SOA\$date\record_spaj_soa_$((Get-Date).ToString("MM")).txt"  

   # Membaca file yang berisi nama file yang sudah disalin sebelumnya
     $record_soa = @()
     if (!(Test-Path -Path $file_record_soa)) {
            New-Item -Path $file_record_soa -ItemType File 
            $record_soa = Get-Content -Path $file_record_soa
        } else { 
             $record_soa = Get-Content -Path $file_record_soa
           }

   # cek file yang belum tercopy dengan rentang waktu h-2 dari tanggal hari ini based on $record_soa
     for($i=1; $i -le 2; $i++) {
          $sync_date = (Get-Date).AddDays(-$i).ToString('dd') 
          $sync_folderspaj = Get-ChildItem "$path\01_SPAJ_SOA\$date\$sync_date\$depart\$area\$loc" | Select-Object FullName
           # cek jika folder nya tidak ada maka lanjut ke next kondisi loop nya
            if ($sync_folderspaj.Count -eq 0) {
                Write-Host " tidak ada file untuk tanggal $sync_date"
                continue
             }
            # copy file spaj di folder spaj soa sesuai tanggal pada $sync_date dan path $sync_folderspaj 
              foreach ($b in $sync_folderspaj) { 
                    $sync_tmpfilename = Split-Path $b -Leaf
                    $sync_filename = $sync_tmpfilename.Replace('}','')
                       if ($record_soa -match $sync_filename) {
                          Write-Host "File $sync_filename sudah pernah di copy sebelumnya."
                            } else {
                                # cek apakah folder 01_SPAJ_SOA nya sudah ada di sftp atau belum
                                $sync_pathspajsoa1 = "/_SPAJ_ESUBMISSION/test/01_SPAJ_SOA/$date/$sync_date/$depart/$area/$loc"
                                   if (!(Test-SFTPPath -Session $sftpsession -Path $sync_pathspajsoa1)) { 
                                          New-SFTPItem -Session $sftpsession -Path $sync_pathspajsoa1 -ItemType Directory -Recurse
                                       }
                                Set-SFTPItem -SessionId $sftpsession.SessionID -Path $b.FullName -Destination $sync_pathspajsoa1
                                Write-Host "File $sync_filename telah dicopy ke folder soa 01 di sftp rds ."
                    }
           }            
      }  

   # copy spaj ke folder 01_SPAJ_SOA di sftp untuk realtime date 
     foreach ($a in $folderspaj)  { 
      $tmpfilename = Split-Path $a -Leaf
      $filename = $tmpfilename.Replace('}','')
        if ($record_soa -match $filename) {
                Write-Host "File $filename sudah pernah di copy sebelumnya."
            } else {
                Set-SFTPItem -SessionId $sftpsession.SessionID -Path $a.FullName -Destination $pathspajsoa1
                Write-Host "File $filename telah dicopy ke folder soa 01 di sftp rds ."
                $record_soa += "$filename`r`n"
          }
     }
   # Simpan nama file yang sudah disalin ke dalam file record
     Set-Content -Path $file_record_soa -Value ($record_soa -join [Environment]::NewLine) 

######################################################################################################################################

# CREATE FOLDER DI FOLDER 02_SPAJ_PENDING_NBA DI SFTP  
# & COPY FILE DI FOLDER 02_SPAJ_PENDING_NBA DI SFTP KE FTP LALU HAPUS FILE YANG TELAH TERCOPY DI SFTP
   
   # copy file ke ftp lokal lalu hapus file
    # cek dan list folder dan file yang akan di copy
   if (!(Test-SFTPPath -Session $sftpsession -Path $pathspajpending)) { 
     New-SFTPItem -Session $sftpsession -Path $pathspajpending -ItemType Directory -Recurse
     $filespajpending = Get-SFTPChildItem -Session $sftpsession -Path $pathspajpending
     $localpathspajpending = "$path\02_SPAJ_PENDING_NBA\$childpath"
        #COPY FILE DARI FOLDER 02_SPAJ_PENDING_NBA 
           foreach ($c in $filespajpending) { 
             Set-SFTPItem -SessionId $sftpsession.SessionID -Path $c.FullName -Destination $localpathspajpending 
                }
           } else {
                $filespajpending = Get-SFTPChildItem -Session $sftpsession -Path $pathspajpending
                $localpathspajpending = "$path\02_SPAJ_PENDING_NBA\$date\$day\$depart\$area\$loc"
                   # COPY FILE DARI FOLDER 02_SPAJ_PENDING_NBA 
                     foreach ($c in $filespajpending) { 
                     Get-SFTPItem -SessionId $sftpsession.SessionID -Path $c.FullName -Destination $localpathspajpending 
                         # Check jika file sudah ada di lokal dan jika sudah ada maka hapus yang di sftp
                            $check = Join-Path -Path $localpathspajpending -ChildPath $c.name
                               if (Test-Path "$check") {
                                  Remove-SFTPItem -SessionId $sftpsession.SessionID -Path $c.FullName -Force
                                  Write-Host "File $($c.name) copied and removed from SFTP."
                               } else {
                                    Write-Host "Error copying file $($c.name) from SFTP to local."
                        }
                }
        }

######################################################################################################################################

# CREATE FOLDER DI FOLDER 03_SPAJ_PROCESS_UW DI SFTP & COPY FILE DI FOLDER 03_SPAJ_PROCESS_UW DI SFTP KE FTP 

   # cek apakah folder sudah terbentuk atau blm
    $pathspajprocessuw | ForEach-Object { 
      if ($_ -and !(Test-SFTPPath -Session $sftpsession -Path $_)) { 
        # buat folder
        New-SFTPItem -Session $sftpsession -Path $_ -ItemType Directory -Recurse
      }
    }
   # copy file ke ftp lokal lalu hapus file
    # list file yang akan di copy
    $filespajprocessuw1 = Get-SFTPChildItem -Session $sftpsession -Path $pathspajprocessuw1
    $filespajprocessuw2 = Get-SFTPChildItem -Session $sftpsession -Path $pathspajprocessuw2
    
    # path spajprocessuw di ftp 
    $localpathspajprocessuw1 = "$path\03_SPAJ_PROCESS_UW\$date\$day\Agency"
    $localpathspajprocessuw2 = "$path\03_SPAJ_PROCESS_UW\$date\$day\Inbranch"

   #COPY FILE DARI FOLDER 03_SPAJ_PROCESS_UW 
    foreach ($e in $filespajprocessuw1) { 
         Set-SFTPItem -SessionId $sftpsession.SessionID -Path $e.FullName -Destination $localpathspajprocessuw1
         # Check jika file sudah ada di lokal dan jika sudah ada maka hapus yang di sftp
             $check2 = Join-Path -Path $localpathspajprocessuw1 -ChildPath $e.name
             if (Test-Path "$check2") {
                  Remove-SFTPItem -SessionId $sftpsession.SessionID -Path $e.FullName -Force
                  Write-Host "File $($e.name) copied and removed from SFTP."
             } else {
                  Write-Host "Error copying file $($e.name) from SFTP to local."
         }
    }

   #COPY FILE DARI FOLDER 03_SPAJ_PROCESS_UW 
    foreach ($f in $filespajprocessuw2) { 
         Set-SFTPItem -SessionId $sftpsession.SessionID -Path $f.FullName -Destination $localpathspajprocessuw2
         # Check jika file sudah ada di lokal dan jika sudah ada maka hapus yang di sftp
             $check3 = Join-Path -Path $localpathspajprocessuw2 -ChildPath $f.name
             if (Test-Path "$check3") {
                  Remove-SFTPItem -SessionId $sftpsession.SessionID -Path $f.FullName -Force
                  Write-Host "File $($f.name) copied and removed from SFTP."
             } else {
                  Write-Host "Error copying file $($f.name) from SFTP to local."
         }
    }

######################################################################################################################################

# putus session sftp
Remove-SFTPSession -Session $sftpsession

# break untuk proses generated report
Start-Sleep -Seconds 3

######################################################################################################################################

# GENERATED CSV REPORT
    $csv1 = Join-Path -Path $path -Childpath "01_SPAJ_SOA\$date\$day\report_01_SPAJ_SOA_$($reportdate).csv"
    
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

######################################################################################################################################

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

