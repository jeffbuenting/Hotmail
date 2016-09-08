Import-Module c:\Scripts\onedrive\onedrive.psm1

$Token = Grant-OneDriveImplicitAuthorization

$Token | FL *



Close-OneDriveSession