# Invoke-Personio
# Lokale Sitzungen lesen das Personio Client-Secret und Access-Tokens aus dem Windows Credential Manager.
# In Azure Automation Runbooks wird das Client-Secret aus Get-AutomationPSCredential gelesen und pro Job ein frisches Access-Token nur im Arbeitsspeicher gehalten.

# Version 1.7.0 18.03.2026 by Klaus Kupferschmid (tempero.it GmbH & hhpberlin GmbH)

#Requires -Modules @{ ModuleName = 'BetterCredentials'; ModuleVersion = '4.5' }

$script:PersonioSettings = [ordered]@{
    ServiceUserName = "HHPBERLIN_USERMANAGER"
    BaseUri         = "https://api.personio.de/v1/company/employees"
    AuthUri         = "https://api.personio.de/v1/auth?"
    MailDomain      = "hhpberlin.de"
}

function Update-PersonioDerivedSettings {
    $script:servicePERUserName = $script:PersonioSettings.ServiceUserName
    $script:Personio_uri = $script:PersonioSettings.BaseUri.TrimEnd('/')
    $script:Personio_auth_uri = $script:PersonioSettings.AuthUri
    $script:mailDomain = $script:PersonioSettings.MailDomain
    $script:Personio_access_token_1 = $script:servicePERUserName+"_PersonioAccessToken_1"
    $script:Personio_access_token_2 = $script:servicePERUserName+"_PersonioAccessToken_2"
}

function Get-PersonioConfiguration {
    [CmdletBinding()]
    param ()

    return [pscustomobject]@{
        ServiceUserName = $script:PersonioSettings.ServiceUserName
        BaseUri         = $script:PersonioSettings.BaseUri
        AuthUri         = $script:PersonioSettings.AuthUri
        MailDomain      = $script:PersonioSettings.MailDomain
        AccessToken1    = $script:Personio_access_token_1
        AccessToken2    = $script:Personio_access_token_2
    }
}

function Set-PersonioConfiguration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)][string]$ServiceUserName,
        [Parameter(Mandatory=$false)][string]$BaseUri,
        [Parameter(Mandatory=$false)][string]$AuthUri,
        [Parameter(Mandatory=$false)][string]$MailDomain
    )

    if ($PSBoundParameters.ContainsKey('ServiceUserName')) {
        $script:PersonioSettings.ServiceUserName = $ServiceUserName
    }

    if ($PSBoundParameters.ContainsKey('BaseUri')) {
        $script:PersonioSettings.BaseUri = $BaseUri.TrimEnd('/')
    }

    if ($PSBoundParameters.ContainsKey('AuthUri')) {
        $script:PersonioSettings.AuthUri = $AuthUri
    }

    if ($PSBoundParameters.ContainsKey('MailDomain')) {
        $script:PersonioSettings.MailDomain = $MailDomain.TrimStart('@')
    }

    Update-PersonioDerivedSettings
    return Get-PersonioConfiguration
}

Update-PersonioDerivedSettings

function Invoke-Personio {
    <#
     .Synopsis
      Connect to Personio Database by REST-API.
    
     .Description
      Connect Personio Database by REST-API. Returns Employee(s) as PSCustomObject(s)
    
     .Parameter Endpoint
      Endpoint is the Appendix of Rest-URL (https://api.personio.de/v1/company/employees).
    
     .Parameter Method
      Method defines what will be done with the selected Endpont.
      Allowed Strings are GET
    
     .Parameter limit
      Pagination attribute to limit the number of employees returned per page.
    
     .Parameter offset
      Pagination attribute to identify the first item in the collection to return.
    
     .Example
       # Retrieve all Employees from Personio
       Invoke-Personio "?"
    
     .Example
       # Get all Personio-Records by Email-Address.
      Invoke-Personio "?email=c.abel%40hhpberlin.de"

        Status                                                                                                          active standard status
        Position                                                                                                      Position standard position
        Employment type                                                                                               internal standard employment_type
        Hire date                                                                                    2010-03-01T00:00:00+01:00 date     hire_date
        Contract ends                                                                                                          date     contract_end_date
        Termination date                                                                                                       date     termination_date
        Termination type                                                                                                       standard termination_type
        Probation period end                                                                                                   date     probation_period_end
        Created at                                                                                   2020-03-13T14:08:29+01:00 date     created_at
        Office                                                                                     @{type=Office; attributes=} standard office
        Profile Picture                                   https://api.personio.de/v1/company/employees/1807395/profile-picture standard profile_picture
        Personalnummer                                                                                                     123 standard
        hhpberlin Kürzel                                                                                                   ABC standard
        Jobtitel                                                                                                      Jobtitel standard
        Ersteintrittsdatum                                                                                                     date     first_company_entry_date
        Eintrittsdatum (aktueller Beschäftigungszeitraum)                                            2010-03-01T00:00:00+01:00 date     latest_employment_start_date
        Austrittsdatum (aktueller Beschäftigungszeitraum)                                                                      date     latest_employment_end_date

    #>
    [CmdletBinding()]
    Param
    (
        [parameter(Position=0,Mandatory=$true)][String] $endpoint,
        [parameter(Position=1,Mandatory=$false)][String] $method = 'GET',
        [parameter(Position=2,Mandatory=$false)] $body,
        [parameter(Position=3,Mandatory=$false)][int32] $limit = 50, # it is mendatory since 7.11.2024, otherwise you only get the fisrt 50
        [parameter(Position=4,Mandatory=$false)][int32] $offset = 0,  # mendatory if limit is set
        [parameter(Mandatory=$false)][switch] $RawOutput
    )
    $token_credentials = Get-Creds
    if (!$token_credentials -or $token_credentials.Count -eq 0) {
        Write-Host "Fehler: Keine Credentials gefunden" -ForegroundColor Red
        return
    }
    $headers=@{}
    $headers.Add("Accept", "application/json")
    $headers.Add("X-Personio-Partner-ID", $Personio_client_id)
    $headers.Add("X-Personio-App-ID", $servicePERUserName)
    $headers.Add("content-type", "application/json")
    # Kombiniere die beiden Token-Teile (zusammen können sie länger als 256 Zeichen sein)
    $token_str1 = ""
    $token_str2 = ""
    if ($token_credentials[0] -and $token_credentials[0].Password) {
        $token_str1 = Convert-SecureStringToPlainText -SecureString $token_credentials[0].Password
    }
    if ($token_credentials.Count -gt 1 -and $token_credentials[1] -and $token_credentials[1].Password) {
        $token_str2 = Convert-SecureStringToPlainText -SecureString $token_credentials[1].Password
    }
    $bearerToken = $token_str1 + $token_str2
    $headers.Add("Authorization", "Bearer $bearerToken")
    if ($body) {
        if ($body.GetType().Name -eq "Hashtable"){
            $body = $body | ConvertTo-Json
            $body = EscapeNonAscii $body
        }
    }
   Switch -Regex ($endpoint){
    '\?' {$uri = $Personio_uri+"?limit=$limit&offset=$offset&"}
    '\?email' {$uri = $Personio_uri+$endpoint}
    "^\d+$" {$uri = $Personio_uri+"/"+$endpoint}
    "/\d+$" {$uri = $Personio_uri+$endpoint} # match if String starts with / followed by Numbers = ID
    Default {$uri = $Personio_uri+$endpoint}
    }
    if($method -eq "PATCH"){
        $uri = $uri+'/'
    }
    <#elseif ($method -ne '/'){
        $uri = $uri+'?'
    }#>
    #$uri = $uri+$endpoint
    $Error.clear()
    try {
        If ($method -eq "PATCH"){
            $response = Invoke-WebRequest -Uri $uri -Method $method -Headers $headers -Body $body -UseBasicParsing -ErrorAction Stop
        }else{
            $response = Invoke-WebRequest -Uri $uri -Method $method -Headers $headers -UseBasicParsing -ErrorAction Stop
        }
    }
    catch {
        switch -RegEx ($PSItem.Exception.Message) {
        "401"   {
                    Write-Host "personio_token ist falsch WebRequest-Error: 401"
                    if (-not $env_runbook) {
                        Remove-StoredCredentialSafe -Target $Personio_access_token_1
                        Remove-StoredCredentialSafe -Target $Personio_access_token_2
                    }
                    If ($Error_401){
                        $script:Error_401 = $null
                        throw 'Fehler "401 Nicht autorisiert" tritt zum zweiten mal auf.'
                    }else{
                        $Error.clear()
                        $script:Error_401 = $true
                        $null = Get-Creds -renew $true
                        $responseObject = Invoke-Personio -endpoint $endpoint -method $method -body $body -limit $limit -offset $offset -RawOutput:$RawOutput
                        $response = $Null
                    }
                }
        "404"   {
                    Write-Host 'Fehler "404" - Benutzer existiert nicht in der Personio Datenbank' -ForegroundColor "Yellow"
                    $script:Error_401 = $null
                    Break
                }
        Default {
                    Write-Host $_ -ForegroundColor Yellow
                    $script:Error_401 = $null
                }   
        }
    }
    if ($response.Content) {
        # Splitt token and save for next request
        $token = ($response.Headers.authorization).Split(" ")[1]
        if ($token.length -ge 201){
            $token1 = $token.substring(0,200)
            $token2 = $token.substring(200)
        }else{
            $token1 = $token
            $token2 = ""
        }
        if (-not $env_runbook) {
            # API Token might be longer than 200 characters, so it is stored in two local credential entries.
            Remove-StoredCredentialSafe -Target $Personio_access_token_1
            Remove-StoredCredentialSafe -Target $Personio_access_token_2
            Set-Credential -Target $Personio_access_token_1 -Credential (New-Object System.Management.Automation.PSCredential('BearerToken', (ConvertTo-SecureString -String $token1 -AsPlainText -Force))) -Type Generic -Description "Personio AccessToken 1" -Persistence Enterprise >$Null
            if ($token2) {
                Set-Credential -Target $Personio_access_token_2 -Credential (New-Object System.Management.Automation.PSCredential('BearerToken', (ConvertTo-SecureString -String $token2 -AsPlainText -Force))) -Type Generic -Description "Personio AccessToken 2" -Persistence Enterprise >$null
            }
        }
        $token = $Null
        $token1 = $Null
        $token2 = $Null
        # read Data.attributes
        $responseObject = ((ConvertFrom-Json  $response.Content).Data.attributes)
    }
    if($method -eq "GET"){
        $return = @()  # WICHTIG: Array initialisieren!
        
        if($responseObject){  
            if ($responseObject.psobject.properties.value[0].gettype().Name -eq "Int32") {
                if($responseObject.psobject.properties.value[4]){
                    # create $employees-Array cut first and last 3 Object (Respond on "Get-Emplyee" indicator for this is first Array-Object is an Integer)
                        $currentPage = @(
                            $responseObject.psobject.properties.value[4..($responseObject.psobject.properties.value.count -4)] |
                                Where-Object { $_ -and $_ -isnot [bool] }
                        )
                    $return += $currentPage
                    Write-Host "Seite $([math]::Floor($offset / $limit) + 1): $($currentPage.count) Employees geladen, gesamt: $($return.count)" -ForegroundColor "Cyan"
                    
                    # Rekursive Aufrufe - wichtig: Rückgabewert erfassen!
                    $recursiveResults = Invoke-Personio -method 'GET' -limit $limit -offset ($offset+$limit) -endpoint $endpoint -RawOutput
                    if ($recursiveResults) {
                        $return += $recursiveResults
                    }
                }else{
                    # Probably was read the Attributes only
                    $return = (ConvertFrom-Json $response.Content).data
                    $response = $Null
                }
            }else{
                # Single employee response: keep the raw attribute object so it can be converted consistently.
                $return += $responseObject
            }
        }
        
        # Ausgabe nur wenn wir am Ende sind (offset=0 = first call, nur dann wird die Nachricht ausgegeben)
        if ($offset -eq 0 -and $return.count -gt 0) {
            if ($return.count -eq 1) {
                Write-Host "Es wurde folgender Employee ausgelesen:" -ForegroundColor "Green"
            } else {
                Write-Host "$($return.count) Employees wurden insgesamt ausgelesen:" -ForegroundColor "Green"
            }
        }
        
        if ($RawOutput -or $endpoint -eq '/attributes') {
            return $return
        }

        return ConvertTo-UserObject $return
    }
}
function Get-Employee {
    <#
     .Synopsis
      Retrieve employe Object(s) from Personio
    
     .Description
      OAuth-Connect Personio Database by REST-API. Returns Employee(s) as PSCustomObject(s)
    
     .Parameter Identity
      Accept either email or Personio-ID as a String.
    
     .Parameter email
      Accept email as a String.

     .Parameter id
      Accept Personio-ID as a int.
    
     .Example
       # Retrieve all Employees from Personio
       Get-Employee
       
     .Example
       # Get a Unique Employee Record from Email-Adress.
       Get-Employee c.abel@hhpberlin.de

        id                           : 1234567
        first_name                   : Vorname
        last_name                    : Nachname
        email                        : v.nachname@hhpberlin.de
        status                       : active
        position                     : Position
        supervisor                   : @{type=Employee; attributes=}
        employment_type              : internal
        hire_date                    : 01.03.2010 00:00:00
        contract_end_date            :
        termination_date             :
        termination_type             :
        probation_period_end         :
        created_at                   : 13.03.2020 14:08:29
        office                       : @{type=Office; attributes=}
        profile_picture              : https://api.personio.de/v1/company/employees/1234567/profile-picture
        Personalnummer               : 123
        hhpberlin_Kuerzel            : AAA
        actual_training_end_date     :
        Jobtitel                     : Jobtitel
        expected_training_end_date   :
        training_start_date          :
        work_permit_expiry_date      :
        first_company_entry_date     :
        latest_employment_start_date : 01.03.2010 00:00:00
        latest_employment_end_date   :
     .Example
       # Get a Unique Employee Record from Personio-ID.
       Get-Employee 1234567
    #>
    [CmdletBinding()]
    param (
        [parameter(Position=1,Mandatory=$false,ValueFromPipeline)]$identity,
        [parameter(Position=2,Mandatory=$false)][int] $id,
        [parameter(Position=3,Mandatory=$false)][String] $email
    )
    
    $endpoint = Get-Endpoint 
    if ($identity){
        $endpoint = Get-Endpoint $identity
    }elseif ($id){
        $endpoint = Get-Endpoint $id
    }elseif ($email){
        $endpoint = Get-Endpoint $email
    }
    return Invoke-Personio $endpoint
}

function Set-Employee {
    <#
     .Synopsis
      Set Attribute Value in Personio Databae
    
     .Description
      OAuth-Connect Personio Database by REST-API. Returns Employee. Set new Attribute Value
    
     .Parameter Identity
      Accept either email or Personio-ID as a String.
    
     .Parameter email
      Accept email as a String.

     .Parameter id
      Accept Personio-ID as a int.
    
     .Example
       # Set new Value 'Ja' to the attribute 'mobile_Phone_number!
       Set-Employee -Id 1234567 -attribute mobile_Phone_number -Value 'Ja'
     .Example
       # Set new Value 'Ja' to the attribute 'mobile_Phone_number!  
       Set-Employee t.eins Mobiltelefon Nein
    #>
    [CmdletBinding()]
    param (
        [parameter(Mandatory=$false,ValueFromPipeline)]$identity,
        [parameter(Position=2,Mandatory=$true)][String] $attribute,
        [parameter(Position=3,Mandatory=$true)][String] $value,
        [parameter(Position=4,Mandatory=$false)][int] $id,
        [parameter(Position=1,Mandatory=$false)][String] $email
        
    )
    if ($identity){
        $endpoint = Get-Endpoint $identity
    }elseif ($id){
        $endpoint = Get-Endpoint $id
    }elseif ($email){
        $endpoint = Get-Endpoint $email
    }
    $employee = Invoke-Personio ($endpoint)
    $user = $employee
    $body = New-Body $attribute $value
    if ($employee){
        $success = Invoke-Personio -endpoint (Get-Endpoint $user.id) -method "PATCH" -body $body
       
    }else{
        throw "Benutzer in Personio nicht gefunden."
    }
    return $success
}
function Sync-MobileOwner{
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]$PhoneContracts,
        [Parameter(Mandatory=$false)][scriptblock]$PhoneContractsProvider
    )

    Write-Host "Lese alle aktiven Employees aus Personio aus" -noNewline -ForegroundColor "Green"
    Write-Host "." -noNewline
    try {
        $allEmployees = Get-Employee | Where-Object status -eq "active"
        Write-Host "OK" -ForegroundColor "Green"
    }
    catch {
        throw "Fehler beim Auslesen der Personio-Datenbank."
    }
    Write-Host "Lese alle aktiven Telefonverträge aus Inventory360 aus" -noNewline -ForegroundColor "Green"
    Write-Host "." -noNewline
    if ($PSBoundParameters.ContainsKey('PhoneContracts')) {
        $phoneContracts = @($PhoneContracts) | Where-Object Status -eq "Active"
        Write-Host "OK" -ForegroundColor "Green"
    } else {
        if (-not $PhoneContractsProvider) {
            $readPhoneContractsCommand = Get-Command -Name Read-PhoneContracts -ErrorAction SilentlyContinue
            if (-not $readPhoneContractsCommand) {
                throw "Read-PhoneContracts ist nicht verfuegbar. Uebergib -PhoneContracts oder -PhoneContractsProvider."
            }

            $PhoneContractsProvider = { Read-PhoneContracts }
        }

        try {
            $phoneContracts = & $PhoneContractsProvider | Where-Object Status -eq "Active"
            Write-Host "OK" -ForegroundColor "Green"
        }
        catch {
            throw "Fehler beim Auslesen der Inventory360-Datenbank."
        }
    }
    
    foreach ($employee in $allEmployees) {
        $contract = $phoneContracts | Where-Object username -eq ($employee.email.split('@')[0])
        $employee_Mobiltelefon = ($employee |Where-Object {$_.Mobiltelefon -eq "Ja" -or $_.Mobiltelefon -eq "Nein"}).Mobiltelefon
        if(($contract -and ($employee_Mobiltelefon -eq "Ja")) -or (!$contract -and ($employee_Mobiltelefon -eq "Nein"))){
            Write-Host "$($employee.email): Mobiltelefon Eintrag in Personio stimmt mit AD überein" -ForegroundColor "Green"
        }else{
            if (!$contract) {$employee_Mobiltelefon = "Nein"}
            else{$employee_Mobiltelefon = "Ja"}
            Write-Host "$($employee.email): Mobiltelefon Eintrag in Personio stimmt nicht mit AD überein. Setzte '$($employee_Mobiltelefon)' in Personio" -ForegroundColor "Yellow" -noNewline
            try {
                Write-Host "." -noNewline
                Set-Employee -id $employee.id -attribute "Mobiltelefon" -value $employee_Mobiltelefon -ErrorAction Stop
                Write-Host "OK" -ForegroundColor "Green"
            }
            catch {
                throw "Fehler beim Schreiben in die Personio-Datenbank."
            }
        }
    }
}
function Show-Attributes{
    <#
     .Synopsis
      Shows a list of available attributes with the authentication used.
    
     .Description
      OAuth-Connect Personio Database by REST-API. Returns a List of Attributes as CustomObject
    
     .Example
       # Shows a list of available attributes with the authentication used.
       Show-Attributes

        key                  label                 type     universal_id
        ---                  -----                 ----     ------------
        first_name           First name            standard first_name
        last_name            Last name             standard last_name
        email                Email                 standard email
        status               Status                standard status
        position             Position              standard position
        supervisor           Supervisor            standard supervisor
        employment_type      Employment type       standard employment_type
        hire_date            Hire date             date     hire_date
        contract_end_date    Contract ends         date     contract_end_date
        termination_date     Termination date      date     termination_date
        termination_type     Termination type      standard termination_type
        termination_reason   Termination reason    standard termination_reason
        probation_period_end Probation period end  date     probation_period_end
        created_at           Created at            date     created_at
        last_modified_at     Last modified         date     last_modified_at
        subcompany           Subcompany            standard subcompany
        office               Office                standard office
        department           Department            standard department
        absence_entitlement  Absence entitlement   standard absence_entitlement
        last_working_day     Last day of work      date     last_working_day
        profile_picture      Profile Picture       standard profile_picture
        team                 Team                  standard team
        dynamic_919111       Personalnummer        standard
        dynamic_1003506      Akademischer Grad     standard name_academic_title
        dynamic_1271341      hhpberlin Kürzel      standard
        dynamic_3600510      Ausbildungsende       date     actual_training_end_date
        dynamic_1395613      Jobtitel              standard
        dynamic_1625230      Private E-Mailadresse standard
        dynamic_919131       Beschäftigungsart     list
        dynamic_3600522      Ausbildungsbeginn     date     training_start_date
        dynamic_3600527      Vorsatzwort           list     name_prefix
        dynamic_1285115      Gültig bis            date
        dynamic_3600531      Ersteintrittsdatum    date     first_company_entry_date
        dynamic_9673413      Mobiltelefon          list
    #>
    Invoke-Personio -Endpoint "/attributes"
}
# private functions
function New-Body {
    param (
        [parameter(Position=1,Mandatory=$true,ValueFromPipeline)][String] $attribute,
        [parameter(Position=2,Mandatory=$true)][String] $value
    )
    $label = ($employee.PSobject.Properties.Value | Where-Object label -eq $attribute).label
    if (!$label){$label = (Show-Attributes | Where-Object universal_id -eq $attribute).label}
    if (!$label){$label = (Show-Attributes | Where-Object key -eq $attribute).label}
    if ($label){
        $key = (Show-Attributes | Where-Object label -eq $label).key
        if ($key -match "dynamic_"){
            $body = @{employee  = @{"custom_attributes" = @{$key = $value
                    }}}
        }else{
            $body = @{employee = @{
                $key = $value
            }}
        }
    }else{
        throw "Das Attribut $attribute bzw. die dazugehörige GUID existiert nicht in Personio."
    }
    return $body
}
function ConvertTo-UserObject ([parameter(Position=1,Mandatory=$false,ValueFromPipeline)]$employees){
    $users = @()
    foreach ($employee in $employees) {
        $attributeValues = @()

        if ($employee -and $employee.PSObject.Properties.Name -contains 'employee') {
            $attributeValues = @($employee.employee)
        } elseif ($employee) {
            $attributeValues = @($employee.PSObject.Properties | ForEach-Object Value)
        }

        if ($attributeValues.Count -eq 0) {
            continue
        }

        $user = New-Object System.Object
        foreach ($property in $attributeValues) {
            if (-not ($property.PSObject.Properties.Name -contains 'value' -and $property.PSObject.Properties.Name -contains 'type')) {
                continue
            }

            switch ($property.type) {
                integer { if ($null -ne $property.value){$property.value = [int]$property.value}}
                date { if ($null -ne $property.value){$property.value = Get-Date $property.value}}
            }
            if ($property.universal_id){
                $noteProperty = $property.universal_id
            }else{
                $noteProperty = $property.label.replace(' ','_')
                $noteProperty = $noteProperty.replace('ü','ue')
                $noteProperty = $noteProperty.replace('ä','ae')
                $noteProperty = $noteProperty.replace('ö','oe')
                $noteProperty = $noteProperty.replace('ß','ss')
            }
            if ($null -eq $property.value){
                $user | Add-Member -type NoteProperty -name $noteProperty -Value $null
            }else {
                $user | Add-Member -type NoteProperty -name $noteProperty -Value $property.value
            }
        }
        $users += $user
    }
    return $users
}
function EscapeNonAscii([string] $s)
{
    # wird nur in Powershell 5.1 benutzt
    $sb = New-Object System.Text.StringBuilder;
    for ([int] $i = 0; $i -lt $s.Length; $i++)
    {
        $c = $s[$i];
        if ($c -gt 127)
        {
            $sb = $sb.Append("\u").Append(([int] $c).ToString("X").PadLeft(4, "0"));
        }
        else
        {
            $sb = $sb.Append($c);
        }
    }
    return $sb.ToString()
}
function Get-CustomIdFromAttribute {
    param (
        [parameter(Position=1,Mandatory=$false,ValueFromPipeline)]$attribute
    )
    $attributes = Invoke-Personio -Endpoint "/attributes"
    $key = ($attributes | Where-Object label -eq $attribute).key
    If ($key){
        return $key
    }else{
        throw "Das Attribut $attribute existiert nicht oder ist nicht fuer die API freigegeben."
    }
}
function Get-Endpoint{
    param (
        [parameter(Position=1,Mandatory=$false,ValueFromPipeline)]$par
    )
    if($par){
        if($par.gettype().Name -eq "Object"){
            $par = $par.email
        }
        switch -Regex ($par)  {
            "\A[a-z0-9!#$%&'*+/=?^_`{|}~-]+(?:\.[a-z0-9!#$%&'*+/=?^_`{|}~-]+)*@(?:[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\.)+[a-z0-9](?:[a-z0-9-]*[a-z0-9])?\z" {$endpoint = "?email="+$par.replace('@','%40')} #UPN
            "^[a-zA-Z0-9]+\.[a-zA-Z0-9]+$" {$endpoint = "?email="+$par+'%40'+$mailDomain} # SAMaccountname
            "^\d+$"      {$endpoint = '/'+$par} # ID
            "attributes" {$endpoint = "/attributes"}
            Default {throw "Identity entspricht weder ID noch email."}
        }
    }else{
        $endpoint = '?'
    }
return $endpoint
}
function Convert-SecureStringToPlainText {
    param (
        [Parameter(Mandatory=$false)]
        [Security.SecureString] $SecureString
    )

    if ($null -eq $SecureString) {
        return ""
    }

    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureString)
    try {
        return [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)
    }
    finally {
        if ($bstr -ne [IntPtr]::Zero) {
            [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr)
        }
    }
}
function Initialize-AutomationEnvironment {
    if ($script:automationEnvironmentInitialized) {
        return
    }

    $script:automationEnvironmentInitialized = $true
    $script:env_runbook = $false
    $script:env_runbook_HybridWorker = $false

    try {
        if ($PSPrivateMetadata.JobId) {
            $script:env_runbook = $true
        }
    }
    catch {
        $script:env_runbook = $false
    }

    if ($script:env_runbook -and $env:COMPUTERNAME -and $env:AUTOMATION_WORKER_NAME) {
        if ($env:COMPUTERNAME -eq $env:AUTOMATION_WORKER_NAME) {
            $script:env_runbook_HybridWorker = $true
        }
    }
}
function Get-StoredCredentialSafe {
    param (
        [Parameter(Mandatory=$true)]
        [string] $Target
    )

    try {
        return Find-Credential -Filter $Target | Select-Object -First 1
    }
    catch {
        return $null
    }
}
function Remove-StoredCredentialSafe {
    param (
        [Parameter(Mandatory=$true)]
        [string] $Target
    )

    try {
        Remove-Credential -Target $Target -Type Generic -ErrorAction Stop
    }
    catch {
    }
}
function New-PersonioTokenCredentialObjects {
    param (
        [Parameter(Mandatory=$true)]
        [string] $Token
    )

    $token1 = $Token
    $token2 = ""
    if ($Token.Length -ge 201) {
        $token1 = $Token.Substring(0,200)
        $token2 = $Token.Substring(200)
    }

    $credentials = @(
        New-Object System.Management.Automation.PSCredential('BearerToken', (ConvertTo-SecureString -String $token1 -AsPlainText -Force))
    )

    if ($token2) {
        $credentials += New-Object System.Management.Automation.PSCredential('BearerToken', (ConvertTo-SecureString -String $token2 -AsPlainText -Force))
    }

    return $credentials
}
function Get-PersonioClientCredential {
    Initialize-AutomationEnvironment

    if ($env_runbook) {
        $automationCredentialCommand = Get-Command -Name Get-AutomationPSCredential -ErrorAction SilentlyContinue
        if (-not $automationCredentialCommand) {
            throw "Get-AutomationPSCredential ist in dieser Runbook-Umgebung nicht verfuegbar."
        }

        try {
            $credential = Get-AutomationPSCredential -Name $servicePERUserName -ErrorAction Stop
        }
        catch {
            throw "AutomationPSCredential mit dem Namen $servicePERUserName konnte nicht gelesen werden."
        }

        if (-not $credential) {
            throw "AutomationPSCredential mit dem Namen $servicePERUserName wurde nicht gefunden."
        }

        $script:Personio_client_id = $credential.UserName
        return $credential
    }

    $storedCred = Get-StoredCredentialSafe -Target $servicePERUserName
    if ($storedCred) {
        $script:Personio_client_id = $storedCred.UserName
        return $storedCred
    }

    Write-Host "Personio Client-ID & Secret wurde noch nicht festgelegt!" -ForegroundColor "Yellow"
    $cred = Microsoft.PowerShell.Security\Get-Credential -Message 'Geben Sie "Personio ClientID" als Benutzername und "Secret" als Passwort ein'
    Set-Credential -Target $servicePERUserName -Credential $cred -Type Generic -Persistence Enterprise -Description "Personio Client ID & Secret" >$null
    $script:Personio_client_id = $cred.UserName
    return $cred
}
function Request-PersonioAccessToken {
    param (
        [Parameter(Mandatory=$true)]
        [System.Management.Automation.PSCredential] $Credential
    )

    # In Runbooks the access token is requested for the current job and only returned in memory.

    $script:Personio_client_id = $Credential.UserName
    $headers = @{
        "Accept" = "application/json"
        "X-Personio-App-ID" = $servicePERUserName
    }
    $body = @{
        "client_id" = $Credential.UserName
        "client_secret" = Convert-SecureStringToPlainText -SecureString $Credential.Password
    }

    try {
        $response = Invoke-WebRequest -Uri $Personio_auth_uri -Method POST -Headers $headers -Body $body -UseBasicParsing -ErrorAction Stop
    }
    catch {
        if ($PSItem.Exception.Message -match "403") {
            if (-not $env_runbook) {
                Remove-StoredCredentialSafe -Target $servicePERUserName
            }

            if ($Error_403) {
                $script:Error_403 = $null
                throw 'Fehler 403 tritt zum zweiten Mal auf.'
            }

            $Error.Clear()
            $script:Error_403 = $true
            return Request-PersonioAccessToken -Credential (Get-PersonioClientCredential)
        }

        $script:Error_403 = $null
        if ($PSItem.Exception.Message -match "The response content cannot be parsed") {
            throw "Die Antwort konnte nicht geparst werden. Pruefen Sie die PowerShell-HTTP-Antwortverarbeitung in der aktuellen Umgebung."
        }

        throw
    }

    $script:Error_403 = $null
    if (-not $response.Content) {
        throw "Personio hat kein Access Token zurueckgegeben."
    }

    return (ConvertFrom-Json $response.Content).Data.token
}
function Write-TokenToCredentialManager{
    # This helper is intended for local sessions where client secret and access token are cached locally.
    if (-not $script:Personio_client_id) {
        $null = Get-PersonioClientCredential
    }

    $storedSecret = Get-StoredCredentialSafe -Target $servicePERUserName
    if ($storedSecret -and -not $script:Personio_client_id) {
        $script:Personio_client_id = $storedSecret.UserName
    }

    if ([string]::IsNullOrWhiteSpace($script:Personio_client_id)) {
        throw 'Personio Client-ID konnte nicht ermittelt werden.'
    }

    if (!$storedSecret) {
        # Prompt once locally and persist the client secret for subsequent local sessions.
        Write-Host "Personio Client Secret wird benötigt" -ForegroundColor "Yellow"
        $cred = Microsoft.PowerShell.Security\Get-Credential -UserName $Personio_client_id -Message 'Geben Sie das "Personio Client Secret" ein'
        if (-not $cred) {
            throw 'Personio Client-ID & Secret wurden nicht eingegeben.'
        }
        Set-Credential -Target $servicePERUserName -Credential $cred -Type Generic -Persistence Enterprise -Description "Personio Client ID & Secret" >$null
        $personio_client_secret = $cred.Password
    } else {
        $personio_client_secret = $storedSecret.Password
    } 
    $headers = @{
        "Accept" = "application/json"
        "X-Personio-App-ID" = $servicePERUserName
    }
    $body = @{
        "client_id" = $Personio_client_id
        "client_secret" = Convert-SecureStringToPlainText -SecureString $personio_client_secret
    }
    
    try {
        $response = Invoke-WebRequest -Uri $Personio_auth_uri -Method POST -Headers $headers -body $body -UseBasicParsing -ErrorAction Stop
    }
    catch {
        If ($PSItem.Exception.Message -match "403"){
            Write-Host "personio_client_secret ist falsch WebRequest-Error: 403"
            Remove-StoredCredentialSafe -Target $servicePERUserName
            If ($Error_403){
                $script:Error_403 = $null
                throw 'Fehler 403 tritt zum zweiten Mal auf.'
            }else{
                $Error.clear()
                $script:Error_403 = $true
                Write-TokenToCredentialManager
            }
        }else{
            $script:Error_403 = $null
        }
        If ($PSItem.Exception.Message -match "The response content cannot be parsed"){
            throw "Die Antwort konnte nicht geparst werden. Pruefen Sie die PowerShell-HTTP-Antwortverarbeitung in der aktuellen Umgebung."
        }
    }
    $personio_client_secret = $null
    IF($response.Content){
        $token = (ConvertFrom-Json $response.Content).Data.token
            if ($token.length -ge 201){
                $token1 = $token.substring(0,200)
                $token2 = $token.substring(200)
            }else{
                $token1 = $token
                $token2 = ""
            }
        # Locally the token is split across two credential entries because it may exceed one entry size.
        Set-Credential -Target $Personio_access_token_1 -Credential (New-Object System.Management.Automation.PSCredential('BearerToken', (ConvertTo-SecureString -String $token1 -AsPlainText -Force))) -Type Generic -Description "Personio AccessToken 1" -Persistence Enterprise >$Null
        # Only store the second token part when it exists.
        if ($token2) {
            Set-Credential -Target $Personio_access_token_2 -Credential (New-Object System.Management.Automation.PSCredential('BearerToken', (ConvertTo-SecureString -String $token2 -AsPlainText -Force))) -Type Generic -Description "Personio AccessToken 2" -Persistence Enterprise >$null
        } else {
            Remove-StoredCredentialSafe -Target $Personio_access_token_2
        }
        $response = $Null
        $token1 = $Null
        $token2 = $Null
    }
    $credentials = @()
    $cred1 = Get-StoredCredentialSafe -Target $Personio_access_token_1
    $cred2 = Get-StoredCredentialSafe -Target $Personio_access_token_2
    if ($cred1) { $credentials += $cred1 }
    if ($cred2) { $credentials += $cred2 }
    return $credentials
}
function Get-Creds {
    Param
    (
        [parameter(Position=0,Mandatory=$false,ValueFromPipeline)] $connectionTarget="PER",
        [parameter(Position=1,Mandatory=$false)][boolean] $renew=$false
    )

    Initialize-AutomationEnvironment
    
    switch ($connectionTarget) {
        PER { }
        Default { throw "Internal Error: Get-Creds wurde mit einem unbekannten Dienst-Kuerzel aufgerufen." }
    }

    if ($env_runbook) {
        # Runbooks always obtain a fresh token from the stored client credential and keep it only in memory.
        $clientCredential = Get-PersonioClientCredential
        $token = Request-PersonioAccessToken -Credential $clientCredential
        return @(New-PersonioTokenCredentialObjects -Token $token)
    }

    $credential = @()
    try {
        $cred1 = Get-StoredCredentialSafe -Target $Personio_access_token_1
        $cred2 = Get-StoredCredentialSafe -Target $Personio_access_token_2
        if ($cred1) { $credential += $cred1 }
        if ($cred2) { $credential += $cred2 }
    }
    catch {
        $Error
    }

    if ($renew -or $credential.count -lt 1 -or $null -in $credential) {
        $null = Get-PersonioClientCredential
        $credential = Write-TokenToCredentialManager
    }

    return $credential
} # Import credentials for the current runtime environment
# Public functions
Export-ModuleMember Invoke-Personio
Export-ModuleMember Get-Employee
Export-ModuleMember Set-Employee
Export-ModuleMember Sync-MobileOwner
Export-ModuleMember Show-Attributes
Export-ModuleMember Get-PersonioConfiguration
Export-ModuleMember Set-PersonioConfiguration