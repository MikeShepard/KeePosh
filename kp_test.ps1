Import-Module kp_utils -force

initialize-keepass -PathToKeePassFolder C:\temp\KeePass-2.28\

$kpdb=get-KPDatabase -path 'C:\users\mike_sh\Documents\AccountOnly.kdbx' -UseCurrentAccount 

get-KPPassword -UserName Michael321 -kpDatabase $kpdb

$entry=get-KPAccount -UserName Michael321 -kpDatabase $kpdb

get-KPPassword -Account $entry  

$cred=get-KPAccount -UserName Mi*  -kpDatabase $kpdb -asCredential

$cred 

$acct=get-kpaccount -Title NewEntry -kp $kpdb

set-KPPassword -Account $acct -password Fred -kpDatabase $kpdb