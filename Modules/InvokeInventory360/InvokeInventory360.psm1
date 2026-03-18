# Invoke-Inventory V1.9.6 18.03.2026 by Klaus Kupferschmid
# Lokale Sitzungen lesen das API-Token aus dem Windows Credential Manager.
# In Azure Automation Runbooks wird das API-Token ueber Get-AutomationPSCredential gelesen; es gibt dort keinen lokalen Fallback.

#Requires -Modules ImportExcel, @{ ModuleName = 'BetterCredentials'; ModuleVersion = '4.5' }

$script:Inventory360Settings = [ordered]@{
    ServiceUserName       = "HHPBERLIN_INVENTORY360"
    ApiBaseUri            = "https://hhp.enteksystems.de/api/2.0/"
    DefaultOrganization   = "hhpberlin - Ingenieure fuer Brandschutz GmbH"
    InitialContractNumber = @{
        Telekom  = "TM058-1"
        Vodafone = "V275095-1"
    }
}

function Update-Inventory360DerivedSettings {
    $script:serviceINVUserName = $script:Inventory360Settings.ServiceUserName
    $script:InventoryApiBaseUri = $script:Inventory360Settings.ApiBaseUri
    if (-not $script:InventoryApiBaseUri.EndsWith('/')) {
        $script:InventoryApiBaseUri += '/'
    }
    $script:InventoryDefaultOrganization = $script:Inventory360Settings.DefaultOrganization
    $script:Inventory_token_target = $script:serviceINVUserName+"_ApiToken"
}

function Get-Inventory360Configuration {
    [CmdletBinding()]
    param ()

    return [pscustomobject]@{
        ServiceUserName          = $script:Inventory360Settings.ServiceUserName
        ApiBaseUri               = $script:Inventory360Settings.ApiBaseUri
        DefaultOrganization      = $script:Inventory360Settings.DefaultOrganization
        TelekomInitialContract   = $script:Inventory360Settings.InitialContractNumber.Telekom
        VodafoneInitialContract  = $script:Inventory360Settings.InitialContractNumber.Vodafone
        TokenTarget              = $script:Inventory_token_target
    }
}

function Set-Inventory360Configuration {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)][string]$ServiceUserName,
        [Parameter(Mandatory=$false)][string]$ApiBaseUri,
        [Parameter(Mandatory=$false)][string]$DefaultOrganization,
        [Parameter(Mandatory=$false)][string]$TelekomInitialContract,
        [Parameter(Mandatory=$false)][string]$VodafoneInitialContract
    )

    if ($PSBoundParameters.ContainsKey('ServiceUserName')) {
        $script:Inventory360Settings.ServiceUserName = $ServiceUserName
    }

    if ($PSBoundParameters.ContainsKey('ApiBaseUri')) {
        $script:Inventory360Settings.ApiBaseUri = $ApiBaseUri
    }

    if ($PSBoundParameters.ContainsKey('DefaultOrganization')) {
        $script:Inventory360Settings.DefaultOrganization = $DefaultOrganization
    }

    if ($PSBoundParameters.ContainsKey('TelekomInitialContract')) {
        $script:Inventory360Settings.InitialContractNumber.Telekom = $TelekomInitialContract
    }

    if ($PSBoundParameters.ContainsKey('VodafoneInitialContract')) {
        $script:Inventory360Settings.InitialContractNumber.Vodafone = $VodafoneInitialContract
    }

    Update-Inventory360DerivedSettings
    return Get-Inventory360Configuration
}

Update-Inventory360DerivedSettings


function Invoke-Inventory {
    <#
     .Synopsis
      Connect Inventory360 Database by REST-API.
    
     .Description
      Connect Inventory360 Database by REST-API. Returns InvokeRestMethod-Status
    
     .Parameter Endpoint
      Endpoint is the Appendix of Rest-URL.
    
     .Parameter Method
      Method defines what will be done with the selected Endpont.
      Allowed Strings are GET,PUT,POST
    
     .Parameter Body
      Body defines a content Attribut in JSON-Format. It is Mandatory if Methods PUT and POST are used.  
    
     .Parameter Id
      Can be a ValueFromPipelene as a PSObject or a String which contains the Inventory360-RecordID
    
     .Example
       # Show Healthy and Counter Status from Inventory360 Database.
       Invoke-Inventory
       health                        counters
       ------                        --------
       @{database=ok; filesystem=ok} @{users=225; uploads=15; quota=1}
    
     .Example
       # Get a List of all PhoneContracts.
       Invoke-Inventory -Endpoint 'assets/phonecontracts' -Method 'GET'
    
        id                 : 3
        number             : TM058-1
        description        : Business Mobile M 2. Generation
        phone_organization : Telekom
        phonenumber        : +49 151 10859789
        contract_model     :
        location           :
        branch             :
        building           :
        floor              :
        room               :
        organization       : hhpberlin - Ingenieure für Brandschutz GmbH
        department         :
        project            :
        username           : p.milhahn
        status             : Active
        organization_to    :
        contract_start     : 2019-03-21
        contract_end       : 2022-03-20
        duration_contract  : 24
        duration_update    : 1
        duration_extend    : 12
        duration_cancel    : 3
        last_update        :
        rate_month         : 37,77
        rate_quarter       :
        rate_year          :
        rate_cycle         : 1
        comment            :
        url                :
        timestamp          : 2021-04-22 23:29:00
    
     .Example
       # Change one or all specific Attributes defined in Jason Body from Id 34
       Invoke-Inventory -Endpoint 'assets/hardware' -Method 'POST' -Body $bodyObj -Id 34
    
     .Example
       # Create a new Hardware Record in Inventory 360 Database
       Invoke-Inventory -Endpoint 'assets/hardware' -Method 'POST' -Body $bodyObj
     .Example
       # Sync Azure Users with Inventory 360 Database
       Invoke-Inventory -Endpoint "admin/users/sync" -Method "POST"
    #>
    [CmdletBinding()]
    Param
    (
        [parameter(Position=0,Mandatory=$false)][String] $Endpoint = 'status',
        [parameter(Position=1,Mandatory=$false)][String] $Method = 'GET',
        [parameter(Position=2,Mandatory=$false)][PSObject] $Body,
        [parameter(Position=3,Mandatory=$false,ValueFromPipeline)][String] $Id
    )
    $token       = Get-InventoryToken
    if (!$token) {
        Write-Host "Fehler: Kein Inventory360 API-Token verfügbar" -ForegroundColor Red
        return $false
    }
    $uri         = $script:InventoryApiBaseUri
    $contentType = 'application/json'
    $headers     = @{'Authorization' = $token;"Accept" = $contentType}
    <#
    If ($PSVersionTable.PSVersion.Major -gt 5){
        #Powershell 7.1
        $bodyJson = ConvertTo-Json $Body -EscapeHandling EscapeNonAscii
    } else {
    #>
        #Powershell 5.1
        $bodyJson = ConvertTo-Json $Body
        $bodyJson = EscapeNonAscii $bodyJson
    #}
    if ($id) {
        $endpoint = $endpoint+'/'+$id
    }
    if ($method -eq 'GET' -or $Endpoint -eq "admin/users/sync") {
        try {
            $stat = Invoke-RestMethod -Uri ($uri+$endpoint) -Headers $headers -ContentType $contentType -Method $method
        }
        catch {
            if ($error[0].Exception.Response.StatusCode -eq "429"){
            Write-Host "Inventory hat derzeit zuviele Anfragen ... warte 5 Sec und versuche es ein 2. mal" -ForegroundColor "Yellow"
            start-sleep -Seconds 5
            $stat = Invoke-RestMethod -Uri ($uri+$endpoint) -Headers $headers -ContentType $contentType -Method $method
            }
            if ($error[0].Exception.Response.StatusCode -eq "429"){
                throw "HTTP 429: The request limit of a user backend has been exceeded. Please try again later."
            }
        }
        
        return $stat    
    } else {
        $error.clear()
        try {
            $stat = Invoke-RestMethod -Uri ($uri+$endpoint) -Headers $headers -ContentType $contentType -Method $method -Body $bodyJson -ErrorVariable RestError -ErrorAction SilentlyContinue 
        }
        catch {
            if ($restError.Errorrecord.ErrorDetails.Message){
                Write-Host "Verbindung zu Inventory war nicht erfolgreich!" -ForegroundColor Red
                Write-Host "Inventory-Rückmeldung:"($restError.Errorrecord.ErrorDetails.Message.split('"')[7]) -ForegroundColor Red
            }
        }
        $success = $true
        If ($stat -match "status=success") {
            $inventoryObject = Invoke-RestMethod -Uri ($uri+$endpoint) -Headers $headers -ContentType $contentType -Method 'GET'
            if($inventoryObject){
                if ($inventoryObject.GetType().Name -eq "String"){
                    $inventoryObject = ConvertFrom-Json ($inventoryObject.replace('imei','IMEI2').replace('eID','eID2'))
                } else{
                    $Members = $Body | Get-Member -MemberType NoteProperty | Select-Object -ExpandProperty Name
                    foreach ($Member in $Members){
                        # Matching Matrix between Propertyname difference in POST and GET Commands
                        switch ($Member) {
                            "Owner" {$testProperty = "username" }
                            "provider" {$testProperty = "phone_organization" }
                            "contract" {$testProperty = "contractnumber" }
                            Default {$testProperty = $Member}
                        }
                        $value = $Body.$member
                        if($value -eq ""){$value = $null}
                        If($inventoryObject | Where-Object {
                            ($_.$testProperty -eq $value) -or ((($_.$testProperty -eq "") -or ($null -eq $_.$testProperty)) -and (($value -eq "") -or ($null -eq $value)))
                        }){
                            Write-Host "Inventory-Eintrag ""$testProperty"" wurde erfolgreich mit dem Wert: ""$value"" beschrieben" -ForegroundColor Green
                            
                        } else {
                            Write-Host "Error: Inventory-Eintrag ""$testProperty"" wurde nicht geschrieben! - aktueller Wert: ""$($inventoryObject.$testProperty)"" " -ForegroundColor Red
                            $success = $false
                        }
                    }
                }
                if($success){
                    return $stat
                }
            }else{
                Write-Host "Error: Inventory-Objekt wurde nicht geschrieben!" -ForegroundColor Red
                $success = $false
            }
        }else{
            $success = $false
        }
    }
    return $success
}
function New-Hardware {
    [CmdletBinding()]
    param (
        [parameter(Position=0,Mandatory=$true,ParameterSetName="SingleValue")][String] $Type, # The hardware type
        [parameter(Position=1,Mandatory=$true,ParameterSetName="SingleValue")][ValidateSet("production", "stock", "deleted", "ordered", "leasingend", "maintenance", "cleaned")][String] $Status, # The status: only allowed values! 
        [parameter(Position=2,Mandatory=$true,ParameterSetName="SingleValue")][String] $Organization, # The organization
        [parameter(Position=3,Mandatory=$true,ParameterSetName="SingleValue")][String] $Name, # The hostname
        [parameter(Position=4,Mandatory=$true,ParameterSetName="SingleValue")][String] $Serialnumber, # The serialNumber
        [parameter(Position=5,Mandatory=$true,ParameterSetName="SingleValue")][ValidateSet("bill", "advance", "after", "creditcard", "leasing", "debit", "paypal")][String] $Payment_method, # The payment method: only allowed values!
        [parameter(Position=6,Mandatory=$false,ParameterSetName="SingleValue")][String] $Location = "ParamNotUsed", # The location
        [parameter(Position=7,Mandatory=$false,ParameterSetName="SingleValue")][String] $Branch = "ParamNotUsed", # The branch
        [parameter(Position=8,Mandatory=$false,ParameterSetName="SingleValue")][String] $Building = "ParamNotUsed", # The building
        [parameter(Position=9,Mandatory=$false,ParameterSetName="SingleValue")][String] $Floor = "ParamNotUsed", # The floor
        [parameter(Position=10,Mandatory=$false,ParameterSetName="SingleValue")][String] $Room = "ParamNotUsed", # The room
        [parameter(Position=11,Mandatory=$false,ParameterSetName="SingleValue")][String] $Department = "ParamNotUsed", # The department
        [parameter(Position=12,Mandatory=$false,ParameterSetName="SingleValue")][String] $Project = "ParamNotUsed", # The project
        [parameter(Position=13,Mandatory=$false,ParameterSetName="SingleValue")][String] $owner = "ParamNotUsed", # The user / owner
        [parameter(Position=14,Mandatory=$false,ParameterSetName="SingleValue")][String] $Phonecontract = "ParamNotUsed", # Assigned phonecontract
        [parameter(Position=15,Mandatory=$false,ParameterSetName="SingleValue")][String] $Manufacturer = "ParamNotUsed", # The manufacturer
        [parameter(Position=16,Mandatory=$false,ParameterSetName="SingleValue")][String] $Model = "ParamNotUsed", # The model
        [parameter(Position=17,Mandatory=$false,ParameterSetName="SingleValue")][String] $Distributor = "ParamNotUsed", # The distributor / provider
        [parameter(Position=18,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Acquisition_value = "-1", # The acquisition cost
        [parameter(Position=19,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Warranty = "-1", # The warranty in months
        [parameter(Position=20,Mandatory=$false,ParameterSetName="SingleValue")][String] $Note = "ParamNotUsed", # Additional notes or comments
        [parameter(Position=21,Mandatory=$false,ParameterSetName="SingleValue")][String] $Os = "ParamNotUsed", # The operatingsystem (Computers, Servers only)
        [parameter(Position=22,Mandatory=$false,ParameterSetName="SingleValue")][String] $Processor = "ParamNotUsed", # The processor model (Computers, Servers only)
        [parameter(Position=23,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Sockets = "-1", # The Number of CPU sockets (Servers only)
        [parameter(Position=24,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Cores = "-1", # The Number of CPU cores (Computers, Servers only)
        [parameter(Position=25,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Frequency = "-1", # The frequency per CPU core in MHz (Computers only)
        [parameter(Position=26,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Ram = "-1", # The amount of RAM in MB (Computers only)
        [parameter(Position=27,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Hdd = "-1", # The total hdd capacity in GB (Computers only)
        [parameter(Position=28,Mandatory=$false,ParameterSetName="SingleValue")][ValidateSet("hdd", "ssd",ParameterSetName="SingleValue")][String] $Hdd_type = "ParamNotUsed", # The hdd type (Computers only): only allowed values! 
        [parameter(Position=29,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Nics = "-1", # The Number of network interfaces (Servers only)
        [parameter(Position=30,Mandatory=$false,ParameterSetName="SingleValue")][String] $Ip_address = "ParamNotUsed", # The primary ip address of the server (Servers only)
        [parameter(Position=31,Mandatory=$false,ParameterSetName="SingleValue")][String] $Ip_netmask = "ParamNotUsed", # The primary subnet mask of the server (Servers only)
        [parameter(Position=32,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Lumen = "-1", # The lumen value (Beamers only)
        [parameter(Position=33,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Ports = "-1", # The Number of network ports (Firewall, Routers, Switches only)
        [parameter(Position=34,Mandatory=$false,ParameterSetName="SingleValue")][Boolean] $Printer, # Has printing capabilities (Printers only)
        [parameter(Position=35,Mandatory=$false,ParameterSetName="SingleValue")][Boolean] $Copy, # Has copy capabilities (Printers only)
        [parameter(Position=36,Mandatory=$false,ParameterSetName="SingleValue")][Boolean] $Fax, # Has fax capabilities (Printers only)
        [parameter(Position=37,Mandatory=$false,ParameterSetName="SingleValue")][Boolean] $Color, # Is a color printer (Printers only)
        [parameter(Position=38,Mandatory=$false,ParameterSetName="SingleValue")][ValidateSet("laser", "inkjet", "thermo",ParameterSetName="SingleValue")][String] $Printer_type = "ParamNotUsed", # The technical printer type (Printers only): only allowed values! 
        [parameter(Position=39,Mandatory=$false,ParameterSetName="SingleValue")][Int] $Size = "-1", # The monitor size in inches (Monitors only)
        [parameter(Position=40,Mandatory=$false,ParameterSetName="SingleValue")][String] $Memory = "ParamNotUsed", # The internal memory / flash size (Mobile devices only)
        [parameter(Position=41,Mandatory=$false,ParameterSetName="SingleValue")][String] $Imei = "ParamNotUsed", # The IMEI Number (Mobile devices only)
        [parameter(Position=28,Mandatory=$true,ValueFromPipeline)][PSObject] $bodyObj
    )
    Process {
        if(!$bodyObj){
            $bodyObj = New-Object System.Object
            If ($type -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Type -Value $type}
            If ($status -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Status -Value $status}
            If ($organization -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Organization -Value $organization}
            If ($name -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Name -Value $name}
            If ($serialnumber -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Serialnumber -Value $serialnumber}
            If ($payment_method -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Payment_method -Value $payment_method}
            If ($location -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Location -Value $location}
            If ($branch -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Branch -Value $branch}
            If ($building -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Building -Value $building}
            If ($floor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Floor -Value $floor}
            If ($room -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Room -Value $room}
            If ($department -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Department -Value $department}
            If ($project -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Project -Value $project}
            If ($owner -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Owner -Value $owner}
            If ($phonecontract -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Phonecontract -Value $phonecontract}
            If ($manufacturer -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Manufacturer -Value $manufacturer}
            If ($model -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Model -Value $model}
            If ($distributor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Distributor -Value $distributor}
            If ($acquisition_value -ne -1){$bodyObj | Add-Member -type NoteProperty -name Acquisition_value -Value $acquisition_value}
            If ($warranty -ne -1 ){$bodyObj | Add-Member -type NoteProperty -name Warranty -Value $warranty}
            If ($note -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Note -Value $note}
            If ($os -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Os -Value $os}
            If ($processor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Processor -Value $processor}
            If ($sockets -ne -1){$bodyObj | Add-Member -type NoteProperty -name Sockets -Value $sockets}
            If ($cores -ne -1){$bodyObj | Add-Member -type NoteProperty -name Cores -Value $cores}
            If ($frequency -ne -1){$bodyObj | Add-Member -type NoteProperty -name Frequency -Value $frequency}
            If ($ram -ne -1){$bodyObj | Add-Member -type NoteProperty -name Ram -Value $ram}
            If ($hdd -ne -1){$bodyObj | Add-Member -type NoteProperty -name Hdd -Value $hdd}
            If ($hdd_type -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Hdd_type -Value $hdd_type}
            If ($nics -ne -1){$bodyObj | Add-Member -type NoteProperty -name Nics -Value $nics}
            If ($ip_address -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Ip_address -Value $ip_address}
            If ($ip_netmask -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Ip_netmask -Value $ip_netmask}
            If ($lumen -ne -1){$bodyObj | Add-Member -type NoteProperty -name Lumen -Value $lumen}
            If ($ports -ne -1){$bodyObj | Add-Member -type NoteProperty -name Ports -Value $ports}
            If ($printer){$bodyObj | Add-Member -type NoteProperty -name Printer -Value $printer}
            If ($copy){$bodyObj | Add-Member -type NoteProperty -name Copy -Value $copy}
            If ($fax){$bodyObj | Add-Member -type NoteProperty -name Fax -Value $fax}
            If ($color){$bodyObj | Add-Member -type NoteProperty -name Color -Value $color}
            If ($printer_type -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Printer_type -Value $printer_type}
            If ($size -ne -1){$bodyObj | Add-Member -type NoteProperty -name Size -Value $size}
            If ($memory -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Memory -Value $memory}
            If ($imei -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name Imei -Value $imei}
        }
        $res = Invoke-Inventory -Endpoint 'assets/hardware' -Method 'POST' -Body $bodyObj
        return $res
    }
}
function New-PhoneContract {
    [CmdletBinding()]
    Param
    (
        [parameter(Position=0,Mandatory=$true,ParameterSetName="PropertyItem")][String] $number, #!Mandatory!
        [parameter(Position=1,Mandatory=$true,ParameterSetName="PropertyItem")][String] $phonenumber, #!Mandatory!
        [parameter(Position=2,Mandatory=$true,ParameterSetName="PropertyItem")][String] $status, #!Mandatory!
        [parameter(Position=3,Mandatory=$true,ParameterSetName="PropertyItem")][String] $organization, #!Mandatory!
        [parameter(Position=4,Mandatory=$true,ParameterSetName="PropertyItem")][String] $contract_start, #!Mandatory!
        [parameter(Position=5,Mandatory=$false,ParameterSetName="PropertyItem")][String] $location = "ParamNotUsed", ##
        [parameter(Position=6,Mandatory=$false,ParameterSetName="PropertyItem")][String] $branch = "ParamNotUsed", ##
        [parameter(Position=7,Mandatory=$false,ParameterSetName="PropertyItem")][String] $building = "ParamNotUsed", ##
        [parameter(Position=8,Mandatory=$false,ParameterSetName="PropertyItem")][String] $floor = "ParamNotUsed", ##
        [parameter(Position=9,Mandatory=$false,ParameterSetName="PropertyItem")][String] $room = "ParamNotUsed", ##
        [parameter(Position=10,Mandatory=$false,ParameterSetName="PropertyItem")][String] $department = "ParamNotUsed", ##
        [parameter(Position=11,Mandatory=$false,ParameterSetName="PropertyItem")][String] $project = "ParamNotUsed", ##
        [parameter(Position=12,Mandatory=$false,ParameterSetName="PropertyItem")][String] $owner = "ParamNotUsed", ##
        [parameter(Position=13,Mandatory=$false,ParameterSetName="PropertyItem")][String] $description = "ParamNotUsed", ##
        [parameter(Position=14,Mandatory=$false,ParameterSetName="PropertyItem")][String] $provider = "ParamNotUsed", ##
        [parameter(Position=15,Mandatory=$false,ParameterSetName="PropertyItem")][String] $organization_to = "ParamNotUsed", ##
        [parameter(Position=16,Mandatory=$false,ParameterSetName="PropertyItem")][String] $contract_end = "ParamNotUsed", ##
        [parameter(Position=17,Mandatory=$false,ParameterSetName="PropertyItem")][String] $duration_contract = "ParamNotUsed", ##
        [parameter(Position=18,Mandatory=$false,ParameterSetName="PropertyItem")][String] $duration_update = "ParamNotUsed", ##
        [parameter(Position=19,Mandatory=$false,ParameterSetName="PropertyItem")][String] $duration_extend = "ParamNotUsed", ##
        [parameter(Position=20,Mandatory=$false,ParameterSetName="PropertyItem")][String] $duration_cancel = "ParamNotUsed", ##
        [parameter(Position=21,Mandatory=$false,ParameterSetName="PropertyItem")][String] $last_update = "ParamNotUsed", ##
        [parameter(Position=22,Mandatory=$false,ParameterSetName="PropertyItem")][String] $rate_month = "ParamNotUsed", ##
        [parameter(Position=23,Mandatory=$false,ParameterSetName="PropertyItem")][String] $rate_quarter = "ParamNotUsed", ##
        [parameter(Position=24,Mandatory=$false,ParameterSetName="PropertyItem")][String] $rate_year = "ParamNotUsed", ##
        [parameter(Position=25,Mandatory=$false,ParameterSetName="PropertyItem")][String] $rate_cycle = "ParamNotUsed", ##
        [parameter(Position=26,Mandatory=$false,ParameterSetName="PropertyItem")][String] $comment = "ParamNotUsed", ##
        [parameter(Position=27,Mandatory=$false,ParameterSetName="PropertyItem")][String] $url = "ParamNotUsed", ##
        [parameter(Mandatory=$true,ValueFromPipeline,ParameterSetName="ObjectItem")][PSObject] $bodyObj
    )
    Process {
        if(!$bodyObj){
            $bodyObj = New-Object System.Object
            $bodyObj | Add-Member -type NoteProperty -name number -Value $number
            $bodyObj | Add-Member -type NoteProperty -name phonenumber -Value $phonenumber
            $bodyObj | Add-Member -type NoteProperty -name status -Value $status
            $bodyObj | Add-Member -type NoteProperty -name organization -Value $organization
            $bodyObj | Add-Member -type NoteProperty -name contract_start -Value $contract_start
            If ($location -ne "ParamNotUsed" -and $location -ne ""){$bodyObj | Add-Member -type NoteProperty -name location -Value $location}
            If ($branch -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name branch -Value $branch}
            If ($building -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name building -Value $building}
            If ($floor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name floor -Value $floor}
            If ($room -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name room -Value $room}
            If ($department -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name department -Value $department}
            If ($project -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name project -Value $project}
            If ($owner -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name owner -Value $owner}
            If ($description -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name description -Value $description}
            If ($provider -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name provider -Value $provider}
            If ($organization_to -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name organization_to -Value $organization_to}
            If ($contract_end -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name contract_end -Value $contract_end}
            If ($duration_contract -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_contract -Value $duration_contract}
            If ($duration_update -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_update -Value $duration_update}
            If ($duration_extend -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_extend -Value $duration_extend}
            If ($duration_cancel -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_cancel -Value $duration_cancel}
            If ($last_update -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name last_update -Value $last_update}
            If ($rate_month -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_month -Value $rate_month}
            If ($rate_quarter -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_quarter -Value $rate_quarter}
            If ($rate_year -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_year -Value $rate_year}
            If ($rate_cycle -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_cycle -Value $rate_cycle}
            If ($comment -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name comment -Value $comment}
            If ($url -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name url -Value $url}
        }
        $res = Invoke-Inventory -Endpoint 'assets/phonecontracts' -Method 'POST' -Body $bodyObj
        return $res
    }
}
function New-SimCard {
    [CmdletBinding()]
    Param
    (
        [parameter(Position=0,Mandatory=$false,ParameterSetName="PropertyItem")][String]$Cardnumber = "ParamNotUsed", ## The cardnumber
        [parameter(Position=1,Mandatory=$true,ParameterSetName="PropertyItem")][String]$contract, # !Mandatory! The assigned contract number
        [parameter(Position=2,Mandatory=$true,ParameterSetName="PropertyItem")][ValidateSet("active", "locked", "disabled", "stored")][String]$Status, # !Mandatory! The SIM card status: only allowed values! 
        [parameter(Position=3,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet(0,1)][Int]$Master = 0, # Is master SIM: only allowed values!
        [parameter(Position=4,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Pin1 = 0, # SIM PIN #1
        [parameter(Position=5,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Pin2 = 0, # SIM PIN #2
        [parameter(Position=6,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Puk1 = 0, # SIM PUK #1
        [parameter(Position=7,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Puk2 = 0, # SIM PUK #2
        [parameter(Position=8,Mandatory=$false,ParameterSetName="PropertyItem")][String]$Comment = "ParamNotUsed", # The contract comment / note
        [parameter(Mandatory=$true,ValueFromPipeline,ParameterSetName="ObjectItem")][PSObject] $bodyObj
    )
    Process {
        if(!$bodyObj){
            $bodyObj = New-Object System.Object
            $bodyObj | Add-Member -type NoteProperty -name contract -Value $contract
            $bodyObj | Add-Member -type NoteProperty -name status -Value $status
            If ($cardnumber -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name cardnumber -Value $cardnumber}
            If ($master -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name master -Value $master}
            If ($pin1 -ne 0){$bodyObj | Add-Member -type NoteProperty -name pin1 -Value $pin1}
            If ($pin2 -ne 0){$bodyObj | Add-Member -type NoteProperty -name pin2 -Value $pin2}
            If ($puk1 -ne 0){$bodyObj | Add-Member -type NoteProperty -name puk1 -Value $puk1}
            If ($puk2 -ne 0){$bodyObj | Add-Member -type NoteProperty -name puk2 -Value $puk2}
            If ($comment -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name comment -Value $comment}
        }
        $res = Invoke-Inventory -Endpoint 'assets/simcards' -Method 'POST' -Body $bodyObj
        return $res
    }
}
function Get-Hardware {
    [CmdletBinding()]
    Param
    (
        [parameter(Position=1,Mandatory=$true,ValueFromPipeline)][String] $Inventorynumber
    )
    $res = Invoke-Inventory -Endpoint 'assets/hardware' -Method 'GET' -Id $inventorynumber
    $res = ConvertFrom-Json ( $res.replace('imei','IMEI2').replace('eID','eID2') )
    return $res
}
function Get-PhoneContract {
    [CmdletBinding()]
    Param
    (
        [parameter(Position=1,Mandatory=$true,ValueFromPipeline)][String] $Id
    )
    $res = Invoke-Inventory -Endpoint 'assets/phonecontracts' -Method 'GET' -Id $id
    return $res
}
function Get-SimCard {
    [CmdletBinding()]
    Param
    (
        [parameter(Position=1,Mandatory=$true,ValueFromPipeline)][String] $Id
    )
    $res = Invoke-Inventory -Endpoint 'assets/simcards' -Method 'GET' -Id $id
    return $res
}
function Read-Contracts {
    [CmdletBinding()]
    $res = Invoke-Inventory -Endpoint 'assets/contracts' -Method 'GET'
    return $res
}
function Read-Leasing {
    [CmdletBinding()]
    $res = Invoke-Inventory -Endpoint 'assets/leasing' -Method 'GET'
    return $res
}
function Read-Hardware {
    [CmdletBinding()]
    $res = Invoke-Inventory -Endpoint 'assets/hardware' -Method 'GET'
    $res = ConvertFrom-Json ( $res.replace('imei','IMEI2').replace('eID','eID2') )
    return $res
}
function Read-PhoneContracts {
    [CmdletBinding()]
    $res = Invoke-Inventory -Endpoint 'assets/phonecontracts' -Method 'GET'
    return $res
}
function Read-SimCards {
    [CmdletBinding()]
    $res = Invoke-Inventory -Endpoint 'assets/simcards' -Method 'GET'
    return $res
}
function Set-Hardware {
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory=$true,ParameterSetName="ObjectItem")]
        [parameter(Position=0,Mandatory=$true,ParameterSetName="PropertyItem")][String] $Inventorynumber, # The inventorynumber
        [parameter(Position=1,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Type = "ParamNotUsed", # The hardware type
        [parameter(Position=2,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet("production", "stock", "deleted", "ordered", "leasingend", "maintenance", "cleaned")][String] $Status = "ParamNotUsed", # The status: only allowed values! 
        [parameter(Position=3,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Organization = "ParamNotUsed", # The organization
        [parameter(Position=4,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Name = "ParamNotUsed", # The hostname
        [parameter(Position=5,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Serialnumber = "ParamNotUsed", # The serialNumber
        [parameter(Position=6,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet("bill", "advance", "after", "creditcard", "leasing", "debit", "paypal")][String] $Payment_method = "ParamNotUsed", # The payment method: only allowed values! 
        [parameter(Position=7,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Location = "ParamNotUsed", # The location
        [parameter(Position=8,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Branch = "ParamNotUsed", # The branch
        [parameter(Position=9,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Building = "ParamNotUsed", # The building
        [parameter(Position=10,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Floor = "ParamNotUsed", # The floor
        [parameter(Position=11,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Room = "ParamNotUsed", # The room
        [parameter(Position=12,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Department = "ParamNotUsed", # The department
        [parameter(Position=13,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Project = "ParamNotUsed", # The project
        [parameter(Position=14,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Owner = "ParamNotUsed", # The user / owner
        [parameter(Position=15,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Phonecontract = "ParamNotUsed", # Assigned phonecontract
        [parameter(Position=16,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Manufacturer = "ParamNotUsed", # The manufacturer
        [parameter(Position=17,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Model = "ParamNotUsed", # The model
        [parameter(Position=18,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Distributor = "ParamNotUsed", # The distributor / provider
        [parameter(Position=19,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Acquisition_value = "-1", # The acquisition cost
        [parameter(Position=20,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Warranty = "-1", # The warranty in months
        [parameter(Position=21,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Note = "ParamNotUsed", # Additional notes or comments
        [parameter(Position=22,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Os = "ParamNotUsed", # The operatingsystem (Computers, Servers only)
        [parameter(Position=23,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Processor = "ParamNotUsed", # The processor model (Computers, Servers only)
        [parameter(Position=24,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Sockets = "-1", # The Number of CPU sockets (Servers only)
        [parameter(Position=25,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Cores = "-1", # The Number of CPU cores (Computers, Servers only)
        [parameter(Position=26,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Frequency = "-1", # The frequency per CPU core in MHz (Computers only)
        [parameter(Position=27,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Ram = "-1", # The amount of RAM in MB (Computers only)
        [parameter(Position=28,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Hdd = "-1", # The total hdd capacity in GB (Computers only)
        [parameter(Position=29,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet("hdd", "ssd")][String] $Hdd_type = "ParamNotUsed", # The hdd type (Computers only): only allowed values! 
        [parameter(Position=30,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Nics = "-1", # The Number of network interfaces (Servers only)
        [parameter(Position=31,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Ip_address = "ParamNotUsed", # The primary ip address of the server (Servers only)
        [parameter(Position=32,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Ip_netmask = "ParamNotUsed", # The primary subnet mask of the server (Servers only)
        [parameter(Position=33,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Lumen = "-1", # The lumen value (Beamers only)
        [parameter(Position=34,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Ports = "-1", # The Number of network ports (Firewall, Routers, Switches only)
        [parameter(Position=35,Mandatory=$false,ParameterSetName="PropertyItem")][Boolean] $Printer, # Has printing capabilities (Printers only)
        [parameter(Position=36,Mandatory=$false,ParameterSetName="PropertyItem")][Boolean] $Copy, # Has copy capabilities (Printers only)
        [parameter(Position=37,Mandatory=$false,ParameterSetName="PropertyItem")][Boolean] $Fax, # Has fax capabilities (Printers only)
        [parameter(Position=38,Mandatory=$false,ParameterSetName="PropertyItem")][Boolean] $Color, # Is a color printer (Printers only)
        [parameter(Position=39,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet("laser", "inkjet", "thermo")][String] $Printer_type = "ParamNotUsed", # The technical printer type (Printers only): only allowed values! 
        [parameter(Position=40,Mandatory=$false,ParameterSetName="PropertyItem")][Int] $Size = "-1", # The monitor size in inches (Monitors only)
        [parameter(Position=41,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Memory = "ParamNotUsed", # The internal memory / flash size (Mobile devices only)
        [parameter(Position=42,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Imei = "ParamNotUsed", # The IMEI Number (Mobile devices only)
        [parameter(Mandatory=$false,ValueFromPipeline,ParameterSetName="ObjectItem")] $pipedObject
    )
    Process {
        if($pipedObject){
            if ($pipedObject.Inventorynumber) {$Inventorynumber = $pipedObject.Inventorynumber}
            if ($pipedObject.GetType().Name -match "Object") {
                $bodyObj = $pipedObject
                $bodyObj.PSObject.Properties.Remove("Inventorynumber")
                # clean Object for later usage for inventory Json-Body
                $illegalMembers = $null
                $illegalMembers = @()
                $illegalMembers += ((Compare-Object (get-help $pscmdlet.commandruntime.tostring()).Syntax.syntaxItem.Parameter.Name $bodyObj.Psobject.Properties.Name) | Where-Object SideIndicator -eq "=>").InputObject.toString()
                foreach ($illegalMember in $illegalMembers){$bodyObj.PSObject.Properties.Remove($illegalMember)}
            }
            if ($pipedObject.GetType().Name -eq "Int64")    {$id = $pipedObject}
            if ($pipedObject.GetType().Name -eq "String")   {[Int64]$id = $pipedObject}
        }
        if (!$bodyObj){
            $bodyObj = New-Object System.Object
            If ($type -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name type -Value $type}
            If ($status -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name status -Value $status}
            If ($organization -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name organization -Value $organization}
            If ($name -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name name -Value $name}
            If ($serialnumber -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name serialnumber -Value $serialnumber}
            If ($payment_method -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name payment_method -Value $payment_method}
            If ($location -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name location -Value $location}
            If ($branch -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name branch -Value $branch}
            If ($building -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name building -Value $building}
            If ($floor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name floor -Value $floor}
            If ($room -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name room -Value $room}
            If ($department -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name department -Value $department}
            If ($project -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name project -Value $project}
            If ($owner -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name owner -Value $owner}
            If ($phonecontract -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name phonecontract -Value $phonecontract}
            If ($manufacturer -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name manufacturer -Value $manufacturer}
            If ($model -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name model -Value $model}
            If ($distributor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name distributor -Value $distributor}
            If ($acquisition_value -ne -1){$bodyObj | Add-Member -type NoteProperty -name acquisition_value -Value $acquisition_value}
            If ($warranty -ne -1){$bodyObj | Add-Member -type NoteProperty -name warranty -Value $warranty}
            If ($note -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name note -Value $note}
            If ($os -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name os -Value $os}
            If ($processor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name processor -Value $processor}
            If ($sockets -ne -1){$bodyObj | Add-Member -type NoteProperty -name sockets -Value $sockets}
            If ($cores -ne -1){$bodyObj | Add-Member -type NoteProperty -name cores -Value $cores}
            If ($frequency -ne -1){$bodyObj | Add-Member -type NoteProperty -name frequency -Value $frequency}
            If ($ram -ne -1){$bodyObj | Add-Member -type NoteProperty -name ram -Value $ram}
            If ($hdd -ne -1){$bodyObj | Add-Member -type NoteProperty -name hdd -Value $hdd}
            If ($hdd_type -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name hdd_type -Value $hdd_type}
            If ($nics -ne -1){$bodyObj | Add-Member -type NoteProperty -name nics -Value $nics}
            If ($ip_address -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name ip_address -Value $ip_address}
            If ($ip_netmask -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name ip_netmask -Value $ip_netmask}
            If ($lumen -ne -1){$bodyObj | Add-Member -type NoteProperty -name lumen -Value $lumen}
            If ($ports -ne -1){$bodyObj | Add-Member -type NoteProperty -name ports -Value $ports}
            If ($printer){$bodyObj | Add-Member -type NoteProperty -name printer -Value $printer}
            If ($copy){$bodyObj | Add-Member -type NoteProperty -name copy -Value $copy}
            If ($fax){$bodyObj | Add-Member -type NoteProperty -name fax -Value $fax}
            If ($color){$bodyObj | Add-Member -type NoteProperty -name color -Value $color}
            If ($printer_type -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name printer_type -Value $printer_type}
            If ($size -ne -1){$bodyObj | Add-Member -type NoteProperty -name size -Value $size}
            If ($memory -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name memory -Value $memory}
            If ($imei -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name imei -Value $imei}
        }
    $res = Invoke-Inventory -Endpoint 'assets/hardware' -Method 'PUT' -Body $bodyObj -Id $Inventorynumber
    return $res
    }
}
function Set-PhoneContract {
    [CmdletBinding()]
    Param
    (
        
        [parameter(Mandatory=$false,ParameterSetName="ObjectItem")]
        [parameter(Position=0,Mandatory=$true,ParameterSetName="PropertyItem")]$Id, # The Identification Number
        [parameter(Position=1,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Number = "ParamNotUsed", # The contract number
        [parameter(Position=2,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Phonenumber = "ParamNotUsed", # The phone number
        [parameter(Position=3,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet("active", "cancelled", "stock")][String] $Status = "ParamNotUsed", # The status: only allowed values! 
        [parameter(Position=4,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Organization = "ParamNotUsed", # The organization
        [parameter(Position=5,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Contract_start = "ParamNotUsed", # The start date
        [parameter(Position=6,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Location = "ParamNotUsed", # The location
        [parameter(Position=7,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Branch = "ParamNotUsed", # The branch
        [parameter(Position=8,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Building = "ParamNotUsed", # The building
        [parameter(Position=9,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Floor = "ParamNotUsed", # The floor
        [parameter(Position=10,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Room= "ParamNotUsed", # The room
        [parameter(Position=11,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Department = "ParamNotUsed", # The department
        [parameter(Position=12,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Project = "ParamNotUsed", # The project
        [parameter(Position=13,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Owner = "ParamNotUsed", # The user / owner
        [parameter(Position=14,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Description = "ParamNotUsed", # The contract description
        [parameter(Position=15,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Provider = "ParamNotUsed", # The contract provider
        [parameter(Position=16,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Organization_to = "ParamNotUsed", # The target organization
        [parameter(Position=17,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Contract_end = "ParamNotUsed", # The end date (either set contract_end OR duration_contract)
        [parameter(Position=18,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Duration_contract = "ParamNotUsed", # The contract duration in months (either set contract_end OR duration_contract)
        [parameter(Position=19,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Duration_update = "ParamNotUsed", # The duration for contract updates in months
        [parameter(Position=20,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Duration_extend = "ParamNotUsed", # The automatic contract extension in months
        [parameter(Position=21,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Duration_cancel = "ParamNotUsed", # The cancellation period in months
        [parameter(Position=22,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Last_update = "ParamNotUsed", # The date of last contract update
        [parameter(Position=23,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Rate_month = "ParamNotUsed", # The monthly contract rate
        [parameter(Position=24,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Rate_quarter = "ParamNotUsed", # The quarterly contract rate
        [parameter(Position=25,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Rate_year = "ParamNotUsed", # The yearly contract rate
        [parameter(Position=26,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Rate_cycle = "ParamNotUsed", # The payment cycle in months
        [parameter(Position=27,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Comment = "ParamNotUsed", # The contract comment / note
        [parameter(Position=28,Mandatory=$false,ParameterSetName="PropertyItem")][String] $Url = "ParamNotUsed", # The link to additonal data / portals
        [parameter(Mandatory=$true,ValueFromPipeline,ParameterSetName="ObjectItem")] $pipedObject
    )
    Process {
        if($pipedObject){
            if ($pipedObject.id) {$id = $pipedObject.id}
            if ($pipedObject.GetType().Name -match "Object") {
                $bodyObj = $pipedObject
                $bodyObj.PSObject.Properties.Remove("Id")
                # clean Object for later usage for inventory Json-Body
                $illegalMembers = $null
                $illegalMembers = @()
                $illegalMembers += ((Compare-Object (get-help $pscmdlet.commandruntime.tostring()).Syntax.syntaxItem.Parameter.Name $bodyObj.Psobject.Properties.Name) | Where-Object SideIndicator -eq "=>").InputObject.toString()
                foreach ($illegalMember in $illegalMembers){$bodyObj.PSObject.Properties.Remove($illegalMember)}
            }
            if ($pipedObject.GetType().Name -eq "Int64") {$id = $pipedObject}
            if ($pipedObject.GetType().Name -eq "String"){[Int64]$id = $pipedObject}
        }
        if (!$bodyObj){
            $bodyObj = New-Object PSCUstomObject
            If ($number -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name number -Value $number}
            If ($phonenumber -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name phonenumber -Value $phonenumber}
            If ($status -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name status -Value $status}
            If ($organization -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name organization -Value $organization}
            If ($contract_start -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name contract_start -Value $contract_start}
            If ($location -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name location -Value $location}
            If ($branch -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name branch -Value $branch}
            If ($building -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name building -Value $building}
            If ($floor -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name floor -Value $floor}
            If ($room -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name room -Value $room}
            If ($department -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name department -Value $department}
            If ($project -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name project -Value $project}
            If ($owner -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name owner -Value $owner}
            If ($description -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name description -Value $description}
            If ($provider -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name provider -Value $provider}
            If ($organization_to -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name organization_to -Value $organization_to}
            If ($contract_end -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name contract_end -Value $contract_end}
            If ($duration_contract -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_contract -Value $duration_contract}
            If ($duration_update -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_update -Value $duration_update}
            If ($duration_extend -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_extend -Value $duration_extend}
            If ($duration_cancel -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name duration_cancel -Value $duration_cancel}
            If ($last_update -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name last_update -Value $last_update}
            If ($rate_month -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_month -Value $rate_month}
            If ($rate_quarter -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_quarter -Value $rate_quarter}
            If ($rate_year -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_year -Value $rate_year}
            If ($rate_cycle -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name rate_cycle -Value $rate_cycle}
            if ($Comment -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name comment -Value $Comment}
            If ($url -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name url -Value $url}
        }
        $res = Invoke-Inventory -Endpoint 'assets/phonecontracts' -Method 'PUT' -Body $bodyObj -Id $id
        return $res
    }
}
function Set-SimCard {
    [CmdletBinding()]
    Param
    (
        [parameter(Mandatory=$false,ParameterSetName="ObjectItem")]
        [parameter(Position=0,Mandatory=$true,ParameterSetName="PropertyItem")]$Id, # The Identification Number
        [parameter(Position=1,Mandatory=$false,ParameterSetName="PropertyItem")][String]$Cardnumber = "ParamNotUsed", # The cardnumber
        [parameter(Position=2,Mandatory=$false,ParameterSetName="PropertyItem")][String]$Contract = "ParamNotUsed", # The assigned contract number
        [parameter(Position=3,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet("active", "locked", "disabled", "stored")][String]$Status = "ParamNotUsed", # The SIM card status: only allowed values! 
        [parameter(Position=4,Mandatory=$false,ParameterSetName="PropertyItem")][ValidateSet(0,1)][Int]$Master, # Is master SIM: only allowed values!
        [parameter(Position=5,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Pin1 = -1, # SIM PIN #1
        [parameter(Position=6,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Pin2 = -1, # SIM PIN #2
        [parameter(Position=7,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Puk1 = -1, # SIM PUK #1
        [parameter(Position=8,Mandatory=$false,ParameterSetName="PropertyItem")][Int]$Puk2 = -1, # SIM PUK #2
        [parameter(Position=9,Mandatory=$false,ParameterSetName="PropertyItem")][String]$Comment = "ParamNotUsed", # The contract comment / note
        [parameter(Mandatory=$true,ValueFromPipeline,ParameterSetName="ObjectItem")] $pipedObject
    )
    Process {
        if($pipedObject){
            if ($pipedObject.id) {$id = $pipedObject.id}
            if ($pipedObject.GetType().Name -match "Object") {
                $bodyObj = $pipedObject
                $bodyObj.PSObject.Properties.Remove("Id")
                # clean Object for later usage for inventory Json-Body
                $illegalMembers = $null
                $illegalMembers = @()
                $illegalMembers += ((Compare-Object (get-help $pscmdlet.commandruntime.tostring()).Syntax.syntaxItem.Parameter.Name $bodyObj.Psobject.Properties.Name) | Where-Object SideIndicator -eq "=>").InputObject.toString()
                foreach ($illegalMember in $illegalMembers){$bodyObj.PSObject.Properties.Remove($illegalMember)}
            }
            if ($pipedObject.GetType().Name -eq "Int64")    {$id = $pipedObject}
            if ($pipedObject.GetType().Name -eq "String")   {[Int64]$id = $pipedObject}
        }
        if (!$bodyObj){
            $bodyObj = New-Object System.Object
            If ($contract -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name contract -Value $contract}
            If ($status -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name status -Value $status}
            If ($cardnumber -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name cardnumber -Value $cardnumber}
            If ($master){$bodyObj | Add-Member -type NoteProperty -name master -Value $master}
            If ($pin1 -ne -1){$bodyObj | Add-Member -type NoteProperty -name pin1 -Value $pin1}
            If ($pin2 -ne -1){$bodyObj | Add-Member -type NoteProperty -name pin2 -Value $pin2}
            If ($puk1 -ne -1){$bodyObj | Add-Member -type NoteProperty -name puk1 -Value $puk1}
            If ($puk2 -ne -1){$bodyObj | Add-Member -type NoteProperty -name puk2 -Value $puk2}
            If ($comment -ne "ParamNotUsed"){$bodyObj | Add-Member -type NoteProperty -name comment -Value $comment}
        }
        $res = Invoke-Inventory -Endpoint 'assets/simcards' -Method 'PUT' -Body $bodyObj -Id $id
        return $res
    }
}
# Special Functions

function Update-PhoneCards([parameter(Position=1,Mandatory=$true,ValueFromPipeline)][String]$csvPath){
    # Import Telekom Report and Update Inventory360
    # Import Telekom Excel-Liste
    $csvs = Import-Csv -Path $csvPath -Encoding "UTF7" -Delimiter ";"
    # Load PhoneContracts
    $phoneContracts = Read-PhoneContracts | Sort-Object -Descending Number
    #$phoneContracts | Format-Table id,number,phonenumber,username,status,owner
    # Load Simcards
    $simCards = Read-SimCards
    #$simCards | Format-Table id,contractnumber,cardnumber,status,comment
    # List Hardware
    #$mobile = Read-Hardware | Where-Object Inventorynumber -like "60????"
    #$mobile | Format-Table Id, hardwaretype,username,phonecontract,imei,manufacturer
    

    foreach($csv in $csvs){
        $rnr =  $csv.Rufnummer.trimstart("49")
        $rnr = "+49 "+$rnr.substring(0,3)+" "+$rnr.substring(3,($rnr.length -3))
        if ($phoneContracts.phonenumber -contains $rnr){
            $phoneContract = $phoneContracts | where-object phonenumber -eq $rnr
            #$csv | Format-List Rufnummer,"Letzte Vertragsverlängerung","Karten-/Profilnummer",Kartentyp,"MultiSIM-Karten-/Profilnummer 1","MultiSIM-Kartentyp 1","MultiSIM-Karten-/Profilnummer 2","MultiSIM-Kartentyp 2"
            #$phoneContract | Format-List PhoneNumber,contract_start
            $sims = $nullread
            $sims = @()
            $sims += $simCards | Where-Object contractnumber -eq $phoneContract.number
            #$sims | FT Id, CardNumber, Comment
            Write-Host "Telefonnummer: " -NoNewline -ForegroundColor Green
            Write-Host $csv.Rufnummer -NoNewline -ForegroundColor White
            Write-Host "   contract.number: " -NoNewline -ForegroundColor Green
            Write-Host ($phoneContract.number) -ForegroundColor White
            if ($csv.'Letzte Vertragsverlängerung'){
                $date = [Datetime]::ParseExact($csv.'Letzte Vertragsverlängerung', 'dd.MM.yyyy', $null)
                $contract_start = $date.GetDateTimeFormats("u").Split(" ")[0]
                if ($contract_start -ne $phoneContract.contract_start){
                    $bodyObj_Contract = New-Object System.Object
                    $bodyObj_Contract | Add-Member -type NoteProperty -name Id -Value $phoneContract.id
                    $bodyObj_Contract | Add-Member -type NoteProperty -name contract_start -Value $contract_start
                    if ($phoneContract.contract_start -ne $bodyObj_Contract.contract_start)
                    {
                        Write-Host "contract_start_alt: "$phoneContract.contract_start" - contract_start_neu: " $bodyObj_Contract.contract_start
                        $bodyObj_Contract | Set-PhoneContract
                    }
                }
            }
            # count CSV-SimCards
            If ($csv."Karten-/Profilnummer"){$simCount =1}
            If ($csv."MultiSIM-Karten-/Profilnummer 1" -ne "\"){$simCount =2}
            If ($csv."MultiSIM-Karten-/Profilnummer 2" -ne "\"){$simCount =3}
            # Set Search-String
            if ($simCount){
                for($i=0; $i -lt $simCount; $i++) {
                    $bodyObj_Sim = New-Object System.Object
                    switch ($i) {
                        "0" {$searchSim = $csv."Karten-/Profilnummer".split("-")[3]}
                        "1" {$searchSim = $csv."MultiSIM-Karten-/Profilnummer 1".split("-")[3]}
                        "2" {$searchSim = $csv."MultiSIM-Karten-/Profilnummer 2".split("-")[3]}
                    }
                    # find matching Sim
                    $sim_match = $null
                    foreach ($sim in $sims){
                        if ($sim.cardnumber -match $searchSim){
                            $sim_match = $sim
                        }
                    }
                    # Change Properties from existing Sim-Card
                    if ($sim_match){ 
                        $change_CNr = $false
                        $change_Com = $false
                        $bodyObj_Sim | Add-Member -type NoteProperty -name Id -Value $sim_match.id
                        if($i -eq 0){
                            if ($sim_match.CardNumber -ne $csv."Karten-/Profilnummer") {
                                $bodyObj_Sim | Add-Member -type NoteProperty -name cardnumber -Value $csv."Karten-/Profilnummer"
                                $bodyObj_Sim | Add-Member -type NoteProperty -name master -Value 1
                                $change_CNr = $true
                            }
                            if ($sim_match.Comment -notlike "*"+$csv.Kartentyp+"*"){
                                $bodyObj_Sim | Add-Member -type NoteProperty -name comment -Value ($sim_match.comment.Split("`n")[0]+"`n"+"Typ: "+$csv.Kartentyp)
                                $change_Com = $true
                            } 
                        }
                        if($i -eq 1){
                            if ($sim_match.CardNumber -ne $csv."MultiSIM-Karten-/Profilnummer 1") {
                                $bodyObj_Sim | Add-Member -type NoteProperty -name cardnumber -Value $csv."MultiSIM-Karten-/Profilnummer 1"
                                $change_CNr = $true
                            }
                            if ($sim_match.Comment -notlike "*"+$csv."MultiSIM-Kartentyp 1"+"*"){
                                $bodyObj_Sim | Add-Member -type NoteProperty -name comment -Value ($sim_match.comment.Split("`n")[0]+"`n"+"Typ: "+$csv."MultiSIM-Kartentyp 1")
                                $change_Com = $true
                            } 
                        }
                        if($i -eq 2){
                            if ($sim_match.CardNumber -ne $csv."MultiSIM-Karten-/Profilnummer 2") {
                                $bodyObj_Sim | Add-Member -type NoteProperty -name cardnumber -Value $csv."MultiSIM-Karten-/Profilnummer 2"
                                $change_CNr = $true
                            }
                            if ($sim_match.Comment -notlike "*"+$csv."MultiSIM-Kartentyp 2"+"*"){
                                $bodyObj_Sim | Add-Member -type NoteProperty -name comment -Value ($sim_match.comment.Split("`n")[0]+"`n"+"Typ: "+$csv."MultiSIM-Kartentyp 2")
                                $change_Com = $true
                            } 
                        }
                        if ($change_CNr){ Write-Host "SimCard gefunden - Alt: "$sim_match.CardNumber" -- Neu:" $bodyObj_Sim.cardnumber}
                        if ($change_Com){ Write-Host $bodyObj_Sim.comment}
                        if ($change_CNr -or $change_Com){
                            $bodyObj_Sim | Set-SimCard
                        } 
                    }else{ # Create new Sim-Card
                        $bodyObj_Sim | Add-Member -type NoteProperty -name contract -Value $phoneContract.number
                        $bodyObj_Sim | Add-Member -type NoteProperty -name status -Value "active"
                        if($i -eq 0){
                            $bodyObj_Sim | Add-Member -type NoteProperty -name cardnumber -Value $csv."Karten-/Profilnummer"
                            $bodyObj_Sim | Add-Member -type NoteProperty -name master -Value 1
                            $bodyObj_Sim | Add-Member -type NoteProperty -name comment -Value ("Ort: `nTyp: "+$csv.Kartentyp)
                        }
                        if($i-eq 1){
                            $bodyObj_Sim | Add-Member -type NoteProperty -name cardnumber -Value $csv."MultiSIM-Karten-/Profilnummer 1"
                            $bodyObj_Sim | Add-Member -type NoteProperty -name comment -Value ("Ort: `nTyp: "+$csv."MultiSIM-Kartentyp 1")
                        }
                        if($i -eq 2){
                            $bodyObj_Sim | Add-Member -type NoteProperty -name cardnumber -Value $csv."MultiSIM-Karten-/Profilnummer 2"
                            $bodyObj_Sim | Add-Member -type NoteProperty -name comment -Value ("Ort: `nTyp: "+$csv."MultiSIM-Kartentyp 2")
                        }
                        Write-Host "SimCard aus Telekom-Report fehlt in Inventory360!" -ForegroundColor Yellow
                        Write-Host "Neu:" $bodyObj_Sim.cardnumber $bodyObj_Sim.comment -ForegroundColor Yellow
                        $bodyObj_Sim | New-SimCard
                    }
                }
            }
        }
    }
}# Import Telekom Report and Update Inventory360
function Pause{
    Write-Host -NoNewLine 'Press any key to continue...';$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
}# Press Key to Continue

function New-PhoneContractFromCsvRecord {
    Param (
        [Parameter(Mandatory=$true,ValueFromPipeline)][PSCustomObject]$csvRecord,
        [parameter(Mandatory=$false)][String]$username,
        [parameter(Mandatory=$false)][String]$PhoneProvider="Telekom"
    )
    $organization = $script:InventoryDefaultOrganization
    $contract_start,$shutdown_Start,$shutdown_End,$contract_end = Get-ContractDatesFromCSVRecord $csvRecord -PhoneProvider $PhoneProvider
    $contract_start = $contract_start.Format
    switch ($PhoneProvider) {
        "Vodafone" {
            $phonenumber = ($csvRecord.TEILNEHMER | Format-PhoneNumber)
            $description = $csvRecord.TARIF
            $comment = $null
        }
        "Telekom" {
            $phonenumber = ($csvRecord.Rufnummer | Format-PhoneNumber)
            $location = ($csvRecord."GP/Wohnort")
            $description = $csvRecord.Tarif
            if ($shutdown_End){
                $comment = "Stilllegung endet am: "+$shutdown_End
            }
            $contract_end = $shutdown_start
        }
        Default { throw "PhoneProvider unbekannt." }
    }
    $contractNumber = Get-NextContractNumber -PhoneProvider $PhoneProvider
    $res = New-PhoneContract `
    -number $contractNumber `
    -phonenumber $phonenumber `
    -status "active"`
    -organization $organization `
    -contract_start $contract_start `
    -location  $location `
    -description $description `
    -provider $PhoneProvider `
    -contract_end $contract_end.Format `
    -comment $comment
    
    try {
        $contractNumber = (Get-PhoneContract -Id $res.id).number
    }
    catch {
        do {
            Start-sleep -Seconds 2
            $contractNumber = (Get-PhoneContract -Id $res.id).number
            $i++
        } until ($res.status -eq "success" -or $i -eq 10)
    }
    if($username){
        Write-Host "Phonecontract wird dem Benutzer "$username" zugewiesen" -ForegroundColor Green
        Set-PhoneContract -Id $res.id -Owner $username
    }
    New-SimCardsFromCsvRecord -csvRecord $csvRecord -contractNumber $contractNumber -PhoneProvider $PhoneProvider
}
function New-SimCardsFromCsvRecord{
    Param (
        [Parameter(Mandatory=$true,ValueFromPipeline)][PSCustomObject]$csvRecord,
        [Parameter(Mandatory=$true)][PSCustomObject]$contractNumber,
        [parameter(Mandatory=$false)][String]$PhoneProvider="Telekom"
    )
    $allSims = Read-SimCards
    if($PhoneProvider -eq "Telekom"){
        $cardNumber = $csvRecord."Karten-/Profilnummer"
        $cardType = $csvRecord.Kartentyp
        $cardPIN = ""
        $cardPUK = ""
        $cardNumber1 = $csvRecord."MultiSIM-Karten-/Profilnummer 1"
        $cardType1 = $csvRecord."MultiSIM-Kartentyp 1"
        $cardPIN1 = ""
        $cardPUK1 = ""
        $cardNumber2 = $csvRecord."MultiSIM-Karten-/Profilnummer 2"
        $cardType2 = $csvRecord."MultiSIM-Kartentyp 2"
        $cardPIN2 = ""
        $cardPUK2 = ""
    }else{
        $cardNumber = $csvRecord.SERIAL_NUMBER
        $cardType = $csvRecord."SIM-Beschreibung"
        if($csvRecord.PIN){[Int]$cardPIN = $csvRecord.PIN.replace(" ","")}
        if($csvRecord.PUK){[Int]$cardPUK = $csvRecord.PUK.replace(" ","")}
        $cardNumber1 = $csvRecord."UC1-SIM"
        $cardType1 = $csvRecord."UC1-Beschreibung"
        if($csvRecord."UC1-PIN"){[Int]$cardPIN1 = $csvRecord."UC1-PIN".replace(" ","")}
        if($csvRecord."UC1-PUK"){[Int]$cardPUK1 = $csvRecord."UC1-PUK".replace(" ","")}
        $cardNumber2 = $csvRecord."UC2-SIM"
        $cardType2 = $csvRecord."UC2-Beschreibung"
        if($csvRecord."UC2-PIN"){[Int]$cardPIN2 = $csvRecord."UC2-PIN".replace(" ","")}
        if($csvRecord."UC2-PUK"){[Int]$cardPUK2 = $csvRecord."UC2-PUK".replace(" ","")}
        

    }
    If ($cardNumber){
        if($allSims.Cardnumber -notcontains $cardNumber){
            New-SimCard `
            -contract $contractNumber `
            -Status "active" `
            -Cardnumber $cardNumber `
            -Master 1 `
            -Pin1 $cardPIN `
            -Puk1 $cardPUK `
            -comment ("Ort: `nTyp: " + $cardType)
        }else{
            Write-Host "SimCard mit der Cardnumber $cardNumber existiert schon." -ForegroundColor Yellow
        }
    }
    If ($cardNumber1){
        if($allSims.Cardnumber -notcontains $cardNumber1){
            New-SimCard `
            -contract $contractNumber `
            -Status "active" `
            -Cardnumber $cardNumber1 `
            -Pin1 $cardPIN1 `
            -Puk1 $cardPUK1 `
            -comment ("Ort: `nTyp: " + $cardType1) `
        }else{
            Write-Host "SimCard mit der Cardnumber $cardNumber1 existiert schon." -ForegroundColor Yellow
        }
    }
    If ($cardNumber2){
        if($allSims.Cardnumber -notcontains $cardNumber2){
            New-SimCard `
            -contract $contractNumber `
            -Status "active" `
            -Cardnumber $cardNumber2 `
            -Pin1 $cardPIN2 `
            -Puk1 $cardPUK2 `
            -comment ("Ort: `nTyp: " + $cardType2) `
        }else{
            Write-Host "SimCard mit der Cardnumber $cardNumber2 existiert schon." -ForegroundColor Yellow
        }
    }
}
function Get-NextContractNumber([parameter(Mandatory=$false)][String]$PhoneProvider="Telekom"){
# this function returns the next free Contractnumber
    $phoneContracts = Read-PhoneContracts
    [Int]$highestContractNumber = 0
    foreach ($phoneContract in $phoneContracts){
        if($phoneContract.number -match '\d'){
            if ([Int]$($phoneContract.number.split("-",2)[1]) -gt $highestContractNumber -and $phoneContract.phone_organization -eq $PhoneProvider){
                $highestContractNumber = [Int]$($phoneContract.number.split("-",2)[1])
                #$latestPhoneContract = $phoneContract
                $ContractNumberPrefix = $phoneContract.number.split("-",2)[0]
            }
        }
    }
    if($highestContractNumber -eq 0 -and $PhoneProvider -eq "Telekom" ){$return =$script:Inventory360Settings.InitialContractNumber.Telekom}
    elseif ($highestContractNumber -eq 0 -and $PhoneProvider -eq "Vodafone") {$return =$script:Inventory360Settings.InitialContractNumber.Vodafone}
    else {$return = $ContractNumberPrefix+"-"+($highestContractNumber +1).toString()}
    return $return
}
function Add-PhoneCardsFromCsvByPhonenumber{
    # create new SimCards to a specific Number from Telekom-Report
    
    Param(
        [parameter(Mandatory=$true,ValueFromPipeline)][String]$Phonenumber,
        [parameter(Mandatory=$true)][String]$csvPath,
        [parameter(Mandatory=$false)][String]$PhoneProvider="Telekom"
    )
    #$csvs = Import-Csv -Path $csvPath -Encoding "UTF7" -Delimiter ";"
    $csvs = Import-Excel -Path $csvPath
    switch ($PhoneProvider) {
        "Vodafone"  {$csvPhonenumber = "TEILNEHMER"
                    $stillegung = "D2-Pause"
                    $contract_start = "CONTRACT_START"
                    $lastContractRenew = "SUB_ACT_DATE"
                    $cardNumber = "SERIAL_NUMBER"
                    $cardNumber1 = "UC1-SIM"
                    $cardNumber2 = "UC2-SIM"
                    $phoneNumberFormat = "E"
        }
        "Telekom"   {$csvPhonenumber = "Rufnummer"
                    $stillegung = "Stillegung"
                    $contract_start = "Vertragsbeginn"
                    $lastContractRenew = "Letzte Vertragsverlängerung"
                    $cardNumber = "Karten-/Profilnummer"
                    $cardNumber1 = "MultiSIM-Karten-/Profilnummer 1"
                    $cardNumber2 = "MultiSIM-Karten-/Profilnummer 2"
                    $phoneNumberFormat = "A"
        }
        Default { throw "PhoneProvider unbekannt." }
    }  
    $csvRecord = $csvs |  Where-Object $csvPhonenumber -eq ($phonenumber| Format-PhoneNumber -Format $phoneNumberFormat)
    if ($csvRecord){
        Write-Host "Eintrag von CSV:" -ForegroundColor Green 
        $csvRecord | Format-Table $csvPhonenumber, $stillegung, $contract_start, $lastContractRenew,$cardNumber,$cardNumber1,$cardNumber2
        Write-Host "Eintrag von Inventory:" -ForegroundColor Green 
        $contract = Read-PhoneContracts | Where-Object phonenumber -eq ($phonenumber| Format-PhoneNumber -Format A)
        if ($contract){
            $contract | Format-Table number, phonenumber, contract_start, contract_end, status, comment
            New-SimCardsFromCsvRecord -csvRecord $csvRecord -contractNumber $contract.Number -PhoneProvider $PhoneProvider
        }else{
            Write-Host "Rufnummer $($phonenumber| Format-PhoneNumber -Format A) wurde in Inventory360 nicht gefunden." -ForegroundColor Yellow
            Write-Host "Neuer PhoneContract wird hinzugefügt" -ForegroundColor Green
            New-PhoneContractFromCsvRecord -csvRecord $csvRecord -PhoneProvider $PhoneProvider
        }
    }else{
        Write-Host "Rufnummer $($phonenumber| Format-PhoneNumber -Format C) wurde in $csvPath nicht gefunden." -ForegroundColor Yellow
    }
    
}
function Format-PhoneNumber{
    # little Helper to Format Phonenumbers
    Param(
        [parameter(Mandatory=$true,ValueFromPipeline)][String]$Phonenumber,
        [parameter(Mandatory=$false)][String]$Format = "A"
    )
    $lean = $Phonenumber.replace('(',"").replace(')',"").replace(" ","").TrimStart("0").TrimStart("+").TrimStart("49")
    switch ($Format) {
        A {$return = "+49 "+$lean.Substring("0",3)+" "+$lean.Substring("3",$lean.length -3)} #For Inventory360
        B {$return = "0"+$lean.Substring("0",3)+" "+$lean.Substring("3",$lean.length -3)} 
        C {$return = "49"+$lean} # for Telekom-Report
        D {$return = "0"+$lean}
        E {$return = $lean}
    }
    return $return
}
function Get-ContractDatesFromCSVRecord {
    Param (
        [Parameter(Mandatory=$true,ValueFromPipeline)][PSCustomObject]$csvRecord,
        [parameter(Mandatory=$false)][String]$PhoneProvider="Telekom"
    )
    switch ($PhoneProvider) {
        "Telekom" {
            if($csvRecord."Letzte Vertragsverlängerung"){
                $contract_start = [Datetime]::ParseExact($csvRecord."Letzte Vertragsverlängerung", 'dd.MM.yyyy', $null)
                $contract_start | Add-Member -Type NoteProperty -Name Format -Value $contract_start.GetDateTimeFormats("u").Split(" ")[0]
            }elseif ($csvRecord.Vertragsbeginn) {
                $contract_start = [Datetime]::ParseExact($csvRecord.Vertragsbeginn, 'dd.MM.yyyy', $null)
                $contract_start | Add-Member -Type NoteProperty -Name Format -Value $contract_start.GetDateTimeFormats("u").Split(" ")[0]
            }
            $shutdown_Start = [Datetime]::ParseExact($csvRecord.Stillegung.replace(" ","").split("-")[0], 'dd.MM.yyyy', $null)
            $shutdown_Start | Add-Member -Type NoteProperty -Name Format -Value $shutdown_Start.GetDateTimeFormats("u").Split(" ")[0]
            $shutdown_End = $csvRecord.Stillegung.replace(" ","").split("-")[1]
        }
        "Vodafone" {
            if($csvRecord."CONTRACT_START"){
                $contract_start = [Datetime]::ParseExact($csvRecord."CONTRACT_START", 'dd.MM.yyyy', $null)
                $contract_start | Add-Member -Type NoteProperty -Name Format -Value $contract_start.GetDateTimeFormats("u").Split(" ")[0]
            }elseif ($csvRecord."SUB_ACT_DATE") {
                $contract_start = [Datetime]::ParseExact($csvRecord."SUB_ACT_DATE", 'dd.MM.yyyy', $null)
                $contract_start | Add-Member -Type NoteProperty -Name Format -Value $contract_start.GetDateTimeFormats("u").Split(" ")[0]
            }
            If($csvRecord."D2-Pause" -ne "N"){
                $shutdown_Start = [Datetime]::ParseExact($csvRecord."D2-Pause", 'dd.MM.yyyy', $null)
                $shutdown_Start | Add-Member -Type NoteProperty -Name Format -Value $shutdown_Start.GetDateTimeFormats("u").Split(" ")[0]
            }else{
                $shutdown_Start = ""
            }
            $shutdown_End = [Datetime]::ParseExact($csvRecord.CONTRACT_END, 'dd.MM.yyyy', $null)
            $shutdown_End | Add-Member -Type NoteProperty -Name Format -Value $shutdown_End.GetDateTimeFormats("u").Split(" ")[0]
            $contract_end = $shutdown_End
        }
        Default { throw "PhoneProvider unbekannt." }
    }
    return $contract_start,$shutdown_Start,$shutdown_End,$contract_end
}
function Compare-ShutdownContracts([parameter(Position=1,Mandatory=$true,ValueFromPipeline)][String]$csvPath){

    $csvs = Import-Csv -Path $csvPath -Encoding "UTF7" -Delimiter ";"
    $shutdown_csvs = $csvs | Where-Object Stillegung -ne ""
    Write-Host "Folgende Verträge sind im Telekom-Report als Stillgelegt vermerkt:" -ForegroundColor Green
    $shutdown_csvs | Format-Table Rufnummer, Stillegung, Vertragsbeginn, "Letzte Vertragsverlängerung" 
    $contracts = Read-PhoneContracts
    Write-Host "Die folgenden PhoneContracts sind in Inventory360 als nicht Aktiv gekennzeichnet:" -ForegroundColor Green
    $notActive_contracts = $contracts | Where-Object Status -ne "active"
    $notActive_contracts | Format-Table number, phonenumber, contract_start, contract_end, status, comment
    foreach ($csvRecord in $shutdown_csvs){
        $contract = $contracts | Where-Object phonenumber -eq ($csvRecord.Rufnummer |Format-PhoneNumber)
        If(!$contract){
            $csvRecord | Format-Table Rufnummer, Stillegung, Vertragsbeginn, "Letzte Vertragsverlängerung"
            Write-Host "Der PhoneContract $($contract.Number) mir Rufnummer: $($contract.phonenumber) konnte nicht in Inventory gefunden werden!" -ForegroundColor Red
            $in = Read-Host -Prompt "Soll ein neuer PhoneContract erstellt werden? (J/N)"
            if ($in -eq "j" -or $in -eq "J"){
                $csvRecord | New-PhoneContractFromCsvRecord
            }else{continue}
        }
        #check if something is different between Telekom-Report and Inventory
        $contract_update = [PSCustomObject]@{Id = $contract.Id;changed = $false} 
        # check if Contract in inventory has Status "cancelled"
        if($contract.status -ne "cancelled"){
            $contract | Format-Table number, phonenumber, contract_start, contract_end, status, comment
            $in = Read-Host -Prompt "PhoneContract $($contract.Number) hat den Status $($contract.status). Soll Status auf 'cancelled' gesetzt werden? (J/N)"
            if ($in -eq "j" -or $in -eq "J"){
                $contract_update | Add-Member -Type NoteProperty -name status -Value "cancelled"
                $contract_update.changed = $true
            }else{continue}
        }
        # set Contract_start
        $contract_start,$shutdown_Start,$shutdown_End = Get-ContractDatesFromCSVRecord $csvRecord
        if ($contract.contract_start -ne $contract_start.Format -and $contract_start) {
            $contract_update | Add-Member -Type NoteProperty -Name contract_start -Value $contract_start.Format
            $contract_update.changed = $true
        }
        if ($contract.contract_end -ne $shutdown_Start.Format) {
            $contract_update | Add-Member -Type NoteProperty -Name contract_end -Value $shutdown_Start.Format
            $contract_update.changed = $true
        }
        if ($contract.comment -ne ("Stilllegung endet am: "+$shutdown_End)) {
            $contract_update | Add-Member -Type NoteProperty -Name comment -Value ("Stilllegung endet am: "+$shutdown_End)
            $contract_update.changed = $true
        }

        if ($contract_update.changed){
            Write-Host "Stilllegung für"$contract.phonenumber" wird aus Telekom-Report übernommen" -ForegroundColor Green
            $contract_update |Format-Table
            # Property die Inventory stört entfernen
            #$contract_update.psobject.properties.remove('changed')
            $contract_update |Set-PhoneContract 
        }else {
            Write-Host "Der PhoneContract $($contract.Number) mit der Rufnummer: $($contract.phonenumber) stimmt mit dem Telekom-Report überein." -ForegroundColor Green
        }
    }
}
function Import-PhoneContracts{
    Param (
        [Parameter(Mandatory=$true,ValueFromPipeline)][PSCustomObject]$csvPath,
        [parameter(Mandatory=$false)][String]$PhoneProvider="Vodafone"
    )
    $csv = Import-Excel -Path $csvPath
    $phoneContracts = Read-PhoneContracts
    $simCards = Read-SimCards
    foreach ($csvRecord in $csv) {
        switch ($PhoneProvider) {
            "Vodafone"  {
                $csvPhonenumber = $csvRecord."TEILNEHMER"
            }
            "Telekom"   {
                $csvPhonenumber = $csvRecord."Rufnummer"  
            }
            Default { throw "PhoneProvider unbekannt." }
        }
    
        $phoneContract = $phoneContracts | Where-Object phonenumber -eq ($csvPhonenumber|Format-PhoneNumber -Format A)
        if($phoneContract){
            Write-Host "Inventory-Eintrag gefunden" -ForegroundColor Green
            $phoneContract | Format-Table 
            Write-Host "Lösche die Telefonnummer $($phoneContract.phonenumber) und ändere den Status auf 'Gekündigt' im alten Telefonvertrag" -ForegroundColor Green
            if($phoneContract.comment){
                $comment = $phoneContract.comment+"`nAlte-Rufnummer: "+$phoneContract.phonenumber
            }else {
                $comment = "Alte-Rufnummer: "+$phoneContract.phonenumber
            }
            Set-PhoneContract -Id $phoneContract.id -Phonenumber "" -status "Cancelled" -Comment $comment
            Write-Host "Lese SIM-Karten aus und setzte sie auf 'Gekündigt'" -ForegroundColor Green
            #$sims = @()
            $sims = $simCards | Where-Object contractnumber -eq $phoneContract.number
            $sims | Format-Table
            $null = $sims | ForEach-Object {Set-SimCard -status "Locked" -id $_.id}
            Write-Host "Erstelle neuen PhoneContract für $($phoneContract.username)" -ForegroundColor Green
            New-PhoneContractFromCsvRecord -csvRecord $csvRecord -Username $phoneContract.username -PhoneProvider "Vodafone" 
        }else{
            Write-Host "Inventory-Eintrag nicht gefunden" -ForegroundColor Yellow
            Write-Host "Lege neue PhonContract in Inventory an"
            New-PhoneContractFromCsvRecord -csvRecord $csvRecord -PhoneProvider "Vodafone" 
        }
    }
}
# private functions
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
function Set-InventoryTokenToCredentialManager {
    # Dieses Verhalten ist nur fuer lokale Sitzungen gedacht.
    $storedCredential = Get-StoredCredentialSafe -Target $Inventory_token_target
    if ($storedCredential) {
        return $storedCredential
    }

    Write-Host "Inventory360 API-Token wird benötigt" -ForegroundColor Yellow
    $credential = Microsoft.PowerShell.Security\Get-Credential -UserName 'InventoryToken' -Message 'Geben Sie den "Inventory360 API-Token" ein'
    Set-Credential -Target $Inventory_token_target -Credential $credential -Type Generic -Persistence Enterprise -Description "Inventory360 API-Token" >$null

    return Get-StoredCredentialSafe -Target $Inventory_token_target
}
function Get-InventoryToken {
    Initialize-AutomationEnvironment

    $credential = $null
    if ($env_runbook) {
        $automationCredentialCommand = Get-Command -Name Get-AutomationPSCredential -ErrorAction SilentlyContinue
        if ($automationCredentialCommand) {
            try {
                $credential = Get-AutomationPSCredential -Name $Inventory_token_target -ErrorAction Stop
            }
            catch {
                Write-Host "AutomationPSCredential mit dem Namen $Inventory_token_target konnte nicht gelesen werden" -ForegroundColor Yellow
            }
        }
        if (!$credential) {
            return $null
        }
    } else {
        $credential = Get-StoredCredentialSafe -Target $Inventory_token_target
        if (!$credential) {
            $credential = Set-InventoryTokenToCredentialManager
        }
    }

    if (!$credential) {
        return $null
    }

    return Convert-SecureStringToPlainText -SecureString $credential.Password
}
function Optimize-Phonenumber {
    $pcs = $phoneContracts | Where-Object phonenumber -notlike "+49 *"
    if ($pcs){
        foreach ($pc in $pcs){
            $pcpnr_old = $pc.phonenumber.toString()
            $pc.phonenumber = $pc.phonenumber.toString().replace("-","")
            $pc.phonenumber = "+49 "+$pc.Phonenumber.substring(0,3)+" "+$pc.Phonenumber.substring(3,($pc.Phonenumber.length -3))
            Write-Host "Ersetze $pcpnr_old gegen "$pc.phonenumber -ForegroundColor Green
            Set-PhoneContract -Id $pc.id -Phonenumber $pc.phonenumber
        }
    }
}# Bring every Phonenumber in PhoneContracts to Format +49 xxx xxxxx

# Public functions
Export-ModuleMember Invoke-Inventory
Export-ModuleMember New-Hardware
Export-ModuleMember New-PhoneContract
Export-ModuleMember New-SimCard
Export-ModuleMember Get-Hardware
Export-ModuleMember Get-PhoneContract
Export-ModuleMember Get-SimCard
Export-ModuleMember Read-Contracts
Export-ModuleMember Read-Leasing
Export-ModuleMember Read-Hardware
Export-ModuleMember Read-PhoneContracts
Export-ModuleMember Read-SimCards
Export-ModuleMember Set-Hardware
Export-ModuleMember Set-PhoneContract
Export-ModuleMember Set-SimCard
Export-ModuleMember Update-PhoneCards
Export-ModuleMember New-PhoneContractFromCsvRecord
Export-ModuleMember New-SimCardsFromCsvRecord
Export-ModuleMember Add-PhoneCardsFromCsvByPhonenumber
Export-ModuleMember Get-NextContractNumber
Export-ModuleMember Format-PhoneNumber
Export-ModuleMember Get-ContractDatesFromCSVRecord
Export-ModuleMember Compare-ShutdownContracts
Export-ModuleMember Import-PhoneContracts
Export-ModuleMember Get-Inventory360Configuration
Export-ModuleMember Set-Inventory360Configuration
