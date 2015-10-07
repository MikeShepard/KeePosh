function Initialize-KeePass{
Param(
        [string]$PathToKeePassFolder = 'C:\temp\KeePass-2.28\'
)

    #Load all .NET binaries in the folder
    (Get-ChildItem -recurse $PathToKeePassFolder|Where-Object {($_.Extension -EQ '.dll') -or ($_.Extension -eq '.exe')} | 
        ForEach-Object { $AssemblyName=$_.FullName; 
                         Try {
                                [Reflection.Assembly]::LoadFile($AssemblyName) } 
                         Catch{ }} ) | out-null
}

function Get-KPDatabase{
    Param(
        [string]$path,
        [string]$KeyFile,
        [string]$MasterPassword,
        [switch]$UseCurrentAccount
    )
    $PwDatabase = new-object KeePassLib.PwDatabase
     $m_pKey = new-object KeePassLib.Keys.CompositeKey
    if($UseCurrentAccount){
        $m_pKey.AddUserKey((New-Object KeePassLib.Keys.KcpUserAccount))
    } 
    if ($KeyFile){
        $m_pKey.AddUserKey((new-object KeePassLib.Keys.KcpKeyFile $KeyFile))
    }
    if ($MasterPassword){
        $m_pKey.AddUserKey((new-object KeePassLib.Keys.KcpPassword $MasterPassword))
    }
    
    $IStatusLogger = New-Object KeePassLib.Interfaces.NullStatusLogger

    $m_ioInfo = New-Object KeePassLib.Serialization.IOConnectionInfo
    $m_ioInfo.Path = $Path
    $PwDatabase.Open($m_ioInfo,$m_pKey,$IStatusLogger)
    $PwDatabase
}

function Get-KPAccount{
    [CmdletBinding()]
    Param([KeePassLib.PwDatabase]$kpDatabase,
          [string]$UserName,
          [string]$Title,
          [string]$Folder,
          [Switch]$AsCredential)
    $pwItems = $kpDatabase.RootGroup.GetObjects($true, $true)
    $found=$false
    foreach($pwItem in $pwItems)
    {
        $EntryUserName=$pwItem.Strings.ReadSafe('UserName')
        $EntryTitle=$pwItem.Strings.ReadSafe('Title')
        $EntryFolder=$pwItem.ParentGroup.Name
        if ((($EntryUserName -like $UserName) -or !$UserName) -and 
           ((($EntryTitle -like $Title) -or !$Title)) -and
           ((($EntryFolder -like $Folder) -or !$Folder)))
        {
            $found=$true
            write-verbose 'Item found'
            if ($AsCredential){
                $pswd=$pwItem.Strings.ReadSafe('Password') | ConvertTo-SecureString -AsPlainText -Force
                new-object System.Management.Automation.PSCredential $EntryUserName,$pswd
            } else {
                $entryURL=$pwItem.Strings.ReadSafe('URL')
                $entryNotes=$pwItem.Strings.ReadSafe('Notes')
                $pwItem | 
                    add-member -MemberType NoteProperty -name UserName -Value $EntryUserName -force -PassThru|
                    add-member -MemberType NoteProperty -name Title -Value $EntryTitle -force -PassThru |
                    add-member -MemberType NoteProperty -name Folder -Value $EntryFolder -force -PassThru | 
                    add-member -MemberType NoteProperty -name URL -Value $EntryURL -force -PassThru | 
                    add-member -MemberType NoteProperty -name Notes -Value $EntryNotes -force -PassThru 
                     

             }
        }
    }
    if (!$found){
        write-Error 'item not found'
    }
  }
function Get-KPPassword{
    [CmdletBinding()]
    Param([KeePassLib.PwDatabase]$kpDatabase,
          [string]$UserName,
          [string]$Title,
          [string]$Folder,
          [Parameter(ParameterSetName='AccountEntry')] $Account)
        
    
    if(!$Account){
        $Account=Get-KPAccount @PSBoundParameters
    }
    switch ($Account.Count) {
        0   {write-error 'Entry not found'}
        1   {$Account.Strings.ReadSafe('Password')}
        else {write-error 'More than one matching row'}
    }


}

function Set-KPPassword{
    [CmdletBinding()]
    Param([KeePassLib.PwDatabase]$kpDatabase,
          [string]$UserName,
          [string]$Title,
          [string]$Password,
          [string]$Folder,
          [Parameter(ParameterSetName='AccountEntry')] $Account)
        
    
    if(!$Account){
        $Account=Get-KPAccount @PSBoundParameters
    }
    switch ($Account.Count) {
        0   {write-error 'Entry not found'}
        1   {   $protectedPassword=new-object KeePassLib.Security.ProtectedString($true,$Password)
            $Account.Strings.Set('Password',$protectedPassword)
            $IStatusLogger = New-Object KeePassLib.Interfaces.NullStatusLogger
            $kpDatabase.Save($IStatusLogger)
        }
        else {write-error 'More than one matching entry'}
    }
}