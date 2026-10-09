$DossierSource = ""
$FichierExcel = ".xlsx"

if (Test-Path $FichierExcel) { Remove-Item $FichierExcel -Force }

$SousDossiers = Get-ChildItem -Path $DossierSource -Directory

Write-Host "Analyse des droits et résolution des groupes AD en cours..." -ForegroundColor Cyan

foreach ($Dossier in $SousDossiers) {
    $NomFeuille = $Dossier.Name
    if ($NomFeuille.Length -gt 31) { $NomFeuille = $NomFeuille.Substring(0, 31) }
    
    Write-Host "Traitement du dossier : $NomFeuille" -ForegroundColor Yellow
    
    $Acl = Get-Acl -Path $Dossier.FullName
    $ResultatsDossier = @()

    foreach ($Rule in $Acl.Access) {
        $Identity = $Rule.IdentityReference.Value
        $Droits = $Rule.FileSystemRights
        $TypeAccess = $Rule.AccessControlType

        $SamAccount = $Identity
        if ($Identity -match "\\") {
            $SamAccount = $Identity.Split("\")[1]
        }

        $AdObject = $null
        try {
            $AdObject = Get-ADObject -Filter "SamAccountName -eq '$SamAccount'" -Properties ObjectClass -ErrorAction SilentlyContinue
        } catch {}

        if ($AdObject -and $AdObject.ObjectClass -eq "group") {
            $Membres = Get-ADGroupMember -Identity $SamAccount -Recursive | Where-Object { $_.objectClass -eq 'user' }
            
            foreach ($Membre in $Membres) {
                $UserProps = Get-ADUser -Identity $Membre.SamAccountName -Properties DisplayName, EmailAddress, Department, Title
                
                $IsDisabled = if ($UserProps.DistinguishedName -match "OU=ArchivedObjects") { "OUI" } else { "NON" }

                $ResultatsDossier += [PSCustomObject]@{
                    Nom         = if ($UserProps.DisplayName) { $UserProps.DisplayName } else { $UserProps.Name }
                    SamAccount  = $UserProps.SamAccountName
                    Email       = $UserProps.EmailAddress
                    Service     = $UserProps.Department
                    Fonction    = $UserProps.Title
                    Disabled    = $IsDisabled
                    TypeAccess  = $TypeAccess
                    Droits      = $Droits
                    Groupe      = $SamAccount
                    ModeAcces   = "Via groupe: $SamAccount"
                }
            }
        }
        elseif ($AdObject -and $AdObject.ObjectClass -eq "user") {
            $UserProps = Get-ADUser -Identity $SamAccount -Properties DisplayName, EmailAddress, Department, Title
            
            $IsDisabled = if ($UserProps.DistinguishedName -match "OU=ArchivedObjects") { "OUI" } else { "NON" }

            $ResultatsDossier += [PSCustomObject]@{
                Nom         = if ($UserProps.DisplayName) { $UserProps.DisplayName } else { $UserProps.Name }
                SamAccount  = $UserProps.SamAccountName
                Email       = $UserProps.EmailAddress
                Service     = $UserProps.Department
                Fonction    = $UserProps.Title
                Disabled    = $IsDisabled
                TypeAccess  = $TypeAccess
                Droits      = $Droits
                Groupe      = ""
                ModeAcces   = "Accès direct"
            }
        }
        else {
            $ResultatsDossier += [PSCustomObject]@{
                Nom         = $Identity
                SamAccount  = $SamAccount
                Email       = ""
                Service     = ""
                Fonction    = ""
                Disabled    = "N/A"
                TypeAccess  = $TypeAccess
                Droits      = $Droits
                Groupe      = $Identity
                ModeAcces   = "Compte local / Système / Non résolu"
            }
        }
    }

    if ($ResultatsDossier.Count -gt 0) {
        $NomTable = "Table_$($NomFeuille -replace '[^a-zA-Z0-9]','_')"
        $ResultatsDossier | Export-Excel -Path $FichierExcel -WorksheetName $NomFeuille -AutoSize -AutoFilter -TableName $NomTable -FreezeTopRow
    }
}

Write-Host "--------------------------------------------------------"
Write-Host "C'est prêt chef : $FichierExcel" -ForegroundColor Green
