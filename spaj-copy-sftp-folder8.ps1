$day = (Get-Date).ToString("dd")
$date = (Get-Date).ToString("yyyyMMdd")
$time = (Get-Date).ToString("yyyyMMdd HH:mm")
$tanggal = (Get-Date).ToString("yyyy-MM") 

$srcpath = "D:\WORK\Task - project\Underwriting-SPAJ\SFTP\08_SPAJ_ACCEPTED"
$destpath = "D:\WORK\Task - project\Underwriting-SPAJ\SFTP\09_SPAJ_ISSUED"

#$srcpath = "\\eli-fs-node01\spaj$\08_SPAJ_ACCEPTED"
#$destpath = "\\eli-fs-node01\spaj$\09_SPAJ_ISSUED"

$no_policy = Import-Csv -Path "D:\WORK\Task - project\Underwriting-SPAJ\Policy_List.csv" -Delimiter "|"  -Header "POLICY_NO" | Select-Object -Skip 1 
$file_policy = Get-ChildItem -Path "D:\WORK\Task - project\Underwriting-SPAJ\SFTP\08_SPAJ_ACCEPTED" | Select-Object -ExpandProperty Name
#$file_policy

foreach ($a in $($no_policy.POLICY_NO)) {
   if ($file_policy -contains $a) { 
    Move-Item -Path "$srcpath\$a" -Destination "$destpath" 
    Add-Content -Path "$srcpath\$($date)_log.log" -Value "file $($a) sudah di pindah ke folder 09_SPAJ_ISSUED pada tanggal $($time) " 
    } else {
    Add-Content -Path "$srcpath\$($date)_log.log" -Value "file $($a) sebelumnya sudah pernah di copy"
    } 
  } 


  # COPY DARI FOLDER NOMOR 9 KE SFTP
 #SFTP SESSION
$sftpdest = "/_SPAJ_ESUBMISSION/test/09_SPAJ_ISSUED/$tanggal/$day"
$destpath = "D:\WORK\Task - project\Underwriting-SPAJ\SFTP\09_SPAJ_ISSUED"

   # authentication
$passwd = ConvertTo-SecureString "password" -AsPlainText -Force
$creds = New-Object System.Management.Automation.PSCredential ("username", $passwd)

    # create session SFTP 
$sftpsession = New-SFTPSession -ComputerName sftp-uat.myequity.id -Credential $creds -AcceptKey

   # cek apakah folder 09_SPAJ_ISSUED nya sudah ada di sftp atau belum
   if (!(Test-SFTPPath -Session $sftpsession -Path $sftpdest)) { 
     New-SFTPItem -Session $sftpsession -Path $sftpdest -ItemType Directory -Recurse
   }

 $spajissued = Get-ChildItem -Path $destpath

   # copy spaj ke folder 01_SPAJ_SOA di sftp
  foreach ($b in $spajissued) { 
    Set-SFTPItem -Sessionid $sftpsession.SessionID -Path $b.FullName -Destination $sftpdest
    Set-SFTPContent -Session $sftpsession -Path "$sftpdest/$($date)_log.log" -Value "file $($b) sudah di pindah ke folder 09_SPAJ_ISSUED pada tanggal $($time) " -Append
   }

   #putus session sftp
Remove-SFTPSession -Session $sftpsession
