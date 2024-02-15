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
$path = "D:\WORK\Task - project\Underwriting-SPAJ\SFTP"
$childpath = "$date\$day\$depart\$area\$loc"
$localpath = "$path\01_SPAJ_SOA\$childpath"

  # path generated report
$pathreport = "$path\01_SPAJ_SOA\$date\$day"

 # SFTP SESSION
   # authentication
$passwd = ConvertTo-SecureString "password" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ("username", $passwd)
   # create session SFTP 
$sftpsession = New-SFTPSession -ComputerName sftp-uat.myequity.id -Credential $creds -AcceptKey

 # Path folder spaj di SFTP
    # path folder 01_SPAJ_SOA
$pathspajsoa1 = "/_SPAJ_ESUBMISSION/test/01_SPAJ_SOA/$date/$day/$depart/$area/$loc"
    # path folder 02_SPAJ_PENDING_NBA
$pathspajpending = "/_SPAJ_ESUBMISSION/test/02_SPAJ_PENDING_NBA/$date/$day/$depart/$area/$loc"
    # path folder 03_SPAJ_PROCESS_UW
$pathspajprocessuw1 = "/_SPAJ_ESUBMISSION/test/03_SPAJ_PROCESS_UW/$date/$day/Agency"
$pathspajprocessuw2 = "/_SPAJ_ESUBMISSION/test/03_SPAJ_PROCESS_UW/$date/$day/Inbranch"
$pathspajprocessuw = @($pathspajprocessuw1,$pathspajprocessuw2)

######################################################################################################################################

# CREATE FOLDER DI FOLDER 02_SPAJ_PENDING_NBA DI SFTP  & COPY FILE DI FOLDER 02_SPAJ_PENDING_NBA DI SFTP KE LOKAL 
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
    
# putus session sftp
Remove-SFTPSession -Session $sftpsession
    }


copySPAJ -depart "Agency" -area 'Area I' -loc "BN Semarang"
copySPAJ -depart "Agency" -area 'Area I' -loc "Indramayu"
copySPAJ -depart "Agency" -area 'Area II' -loc "Denpasar"
