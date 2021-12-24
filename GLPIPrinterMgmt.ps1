# GLPI
$glpiUri = 'YourAddress/apirest.php/'
$yourToken = 'YourToken'

$GlpiSession = [PSCustomObject]@{
    Session      = $null
    Computers    = $null
    Users        = $null
    Documents    = $null
    UserEmail    = $null
    DocumentItem = $null
}

function Start-GLPISession {
    $initSession = $glpiUri + 'initSession'
    $headers = @{
        "Content-Type" = 'application/json';
        "Authorization" = "user_token $yourToken"
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $initSession -UseBasicParsing
    $Session = ($r | ConvertFrom-Json).session_token
    $Script:GlpiSession.Session = $Session
    return $Session
}

function Get-GLPISession{
    if($Script:GlpiSession.Session){
        return $Script:GlpiSession.Session
    }
    Start-GLPISession
}
function Stop-GLPISession {
    param (
        $Session
    )
    if(!$Session){
        $Session = $Script:GlpiSession.Session
    }
    $killSession = $glpiUri + 'killSession'
    $headers = @{
        "Content-Type" = 'application/json';
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $killSession -UseBasicParsing

    $Script:GlpiSession.Session   = $null
    $Script:GlpiSession.Computers = $null
    $Script:GlpiSession.Users     = $null
    $Script:GlpiSession.Documents = $null
    $Script:GlpiSession.UserEmail = $null
    return $r.StatusDescription
}
function Find-GLPIComputer{
    param(
        $name,
        $Session = $Script:GlpiSession.Session
    )

    $type  = 'Computer'
    $uri = $glpiUri + "search/$type" + "?criteria[0][itemtype]=$type&criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]=$name&forcedisplay[0]=2"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)

    Write-Host "INFORMATIONS GLPI:"
    if($result.totalcount -eq 0){
        Write-Host "Ref KO pour $name"
    }else{
        Write-Host $result
        $data = $result.data
        foreach ($ordi in $data){
            if($ordi.1 -eq $name){
                Write-Host $ordi.1
                Write-Host $ordi.70
                Write-Host $ordi.7
                Write-Host $ordi.3
                Write-Host $ordi.5150
                Write-Host $ordi.5197
                Write-Host $ordi.23
                Write-Host $ordi.4
                Write-Host $ordi.40
                Write-Host $ordi.45
                Write-Host
            }
        }
        $data
    }
}

function Get-GLPIComputer{
    param(
        $name,
        $Session = $Script:GlpiSession.Session
    )

    $type  = 'Computer'
    $uri = $glpiUri + "$type" + "/?expand_dropdowns=true&range=0-1000"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)
    $result
}

#Recupere la table des imprimantes
function Get-GLPIPrinterStrict{
    param(
        $Session = (Get-GlpiSession)
    )
    $type = 'Printer'
    Get-GLPIType -Session $Session -Type $type -Fin 10000
}

function Get-GLPIPrinterManufacturer{
    param(
        $Session = (Get-GLPISession)
    )
    (Get-GLPIPrinterStrict -Session $Session).manufacturers_id | Sort-Object -Unique
}

function Get-GLPIPrinterModel{
    param(
        [String]$Manufacturer,
        $Session = (Get-GlpiSession)
    )
    $Printer = (Get-GLPIPrinterStrict -Session $Session)
    if($Manufacturer){
        $Printer = $Printer|?{$_.manufacturers_id -eq $Manufacturer}
    }
    $Printer.printermodels_id | Sort-Object -Unique
}

function Get-GLPIPrinterDriver{
    param(
        [parameter(ValueFromPipelineByPropertyName)]
        [String]$Model,
        $Session = (Get-GLPISession)
    )
    Get-GLPIType -Type PrinterModel -Session $Session | ?{
        $_.name -eq $Model
    } | %{
        if($_.comment -match 'driver\s*:\s*(.+)'){$Matches[1]}
    }
}

#Add driver and code to a printer
#Driver from the printer model
#Code from comment
function Add-GLPIPrinterExtendedInfo{
    param(
        [parameter(ValueFromPipeline)]
        $Printer,
        $Session = (Get-GLPISession)
    )
    Begin{
        $ListPrinter = @()
        $ListPrinterModel = Get-GLPIType -Type PrinterModel -Session $Session 
    }
    Process{
        $ListPrinterModel | ?{
            $_.name -eq $Printer.Model
        } | %{
            if($_.comment -match 'driver\s*:\s*(.+)'){
                Add-Member -InputObject $Printer -MemberType NoteProperty -Name "Driver" -Value ($Matches[1])
            }
            $Printer.Comment | %{
                if(($_ -match 'code\s*:\s*(.+)') -and ($null -eq $Printer.Code)){
                    #On ajoute un seul code!! meme si par erreur plusieurs sont entres dans les commentaires
                    Add-Member -InputObject $Printer -MemberType NoteProperty -Name "Code" -Value ($Matches[1]) 
                }
            }
            $ListPrinter += $Printer
        }
    }
    End{
        $ListPrinter
    }
}


function Get-GLPIPrinter{
    param(
        $Session = (Get-GlpiSession)
    )
    $PrinterList = Get-GLPIPrinterStrict -Session $Session
    $GLPIPrinterList = @()

    if($PrinterList){
        #Recuperation de toutes les interfaces ethernet pour recuperer l'adress mac
        $interfaces = (Get-GLPIType -Session $Session -Type 'NetworkPort' -Fin 100000)
        #Recuperation de toutes les adresses ip des imprimantes
        $IPs = (Get-GLPIType -Session $Session -Type 'IPAddress' -Fin 100000)|?{$_.mainitemtype -eq 'Printer'}       
        #On parcours toutes les imprimantes pour ajouter la mac et l'IP
        foreach($printer in $PrinterList){
            $interface = $interfaces | ?{$_.items_id -eq $printer.name}
            #On verifie si l'IP appartient a l'imprimante
            :boucle foreach ($ip in $IPs){
                $links = $ip.links
                foreach ($link in $links){
                    if($link.rel -eq 'UNKNOWN'){#On utilise un BUG dans l'API de GLPI!
                        if($link.href -match '/(\d+)$'){
                            if($Matches[1] -eq $printer.id){
                                $IPAddress = $ip.name
                                break boucle #On considere naivement que l'imprimante n'a qu'une seule IP
                            }
                        }
                    }
                }
            }
            $GLPIPrinterList += [PSCustomObject]@{
                Id           = $printer.id
                Name         = $printer.name
                Entity       = $printer.entities_id
                Location     = $printer.locations_id
                Manufacturer = $printer.manufacturers_id
                Model        = $printer.printermodels_id
                Serial       = $printer.serial
                Mac          = $interface.mac
                IPAddress    = $IPAddress
                Comment      = $printer.comment
            }
        }
    }
    $GLPIPrinterList
}

function Find-GLPIPrinter{
    param(
        $Session = (Get-GlpiSession),
        $Id,
        [Parameter()]
        [String]$Site,
        [String]$Manufacturer,
        [String]$Model,
        [String]$IPAddress,
        [String]$Mac
    )
    #Info: recherche des locations:
    $printer = Get-GLPIPrinter -Session $Session
    if($printer){
        if($Id){
            $printer = $printer|?{$_.Id -eq $Id}
        }
        if($Site){
            $printer = $printer|?{$_.Location -match $Sites}
        }
        if($Manufacturer){
            $printer = $printer|?{$_.Manufacturer-eq $Manufacturer}
        }
        if($Model){
            $printer = $printer|?{$_.Model -eq $Model}
        }
        if($IPAddress){
            $printer = $printer|?{$_.IPAddress -eq $IPAddress}
        }
        if($Mac){
            $printer = $printer|?{
                foreach($MacAddress in $_.Mac){
                    if(($null -ne $MacAddress) -and
                            ($MacAddress.Replace(':','').Replace('-','').ToLower() -eq
                             $Mac.Replace(':','').Replace('-','').ToLower())){
                        $True
                        break
                    }
                }
                $False
            }
        }

        $printer
    }
}

function Get-GLPINetworkNameIP{
    param(
        $Session = (Get-GlpiSession),
        $NetworkNameId
    )

    if(!$Session){
        $Session = $Script:GlpiSession.Session
    }

    $type = 'NetworkName'
    $uri = $glpiUri + "$type" + "/" + $NetworkNameId + "/IPAddress"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)
    $result
}

function Get-GLPIType{
    param(
        $Session = (Get-GlpiSession),
        $Type = 'Computer',
        $Debut = 0,
        $Fin   = 1000
    )

    $uri = $glpiUri + "$type" + "/?expand_dropdowns=true&range=$Debut-$Fin"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)
    $result
}

function Get-GLPIItem{
    param(
        $Type,
        $Id,
        $Session = (Get-GlpiSession)
    )
    $uri = $glpiUri + "$type" + '/' + "$Id" + "?with_networkports"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing
    ($r | ConvertFrom-Json)
}

function Get-GLPISearchOptions{
    param(
        $Type,
        $Session = (Get-GLPISession)
    )
    $uri = $glpiUri + "listSearchOptions/$Type"

    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }

    Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing | ConvertFrom-Json

}
function Search-GLPIItem{
    param(
        $Type = 'IPAddress',
        $name,
        $Session = (Get-GLPISession)
    )
    # forcedisplay 126 pour l'adresse IP et 2 pour la clef de la table 21 pour l'adresse MAC
    $uri = $glpiUri + "search/$Type" + "?criteria[0][itemtype]=$Type&criteria[0][field]=1&criteria[0][searchtype]=contains&criteria[0][value]=$name&range=0-1000&forcedisplay[0]=126&forcedisplay[1]=2&forcedisplay[2]=21"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)

    Write-Host "INFORMATIONS GLPI:"
    if($result.totalcount -eq 0){
        Write-Host "Ref KO pour $name"
    }else{
        $data = $result.data
        $data
    }
}

function Request-GLPILink{
    param(
        $Link,
        $Session = (Get-GlpiSession)
    )
    $uri = $Link -replace "^http://(\d+)\.(\d+)\.(\d+).(\d+)/", 'http://support.greta06.fr/'
    Write-Host $uri
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)
    $result
}

function Get-GLPIMyEntities{
    param(
        $Session = (Get-GlpiSession)
    )

    $uri = $glpiUri + "getMyEntities/"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)
    $result
}

function Get-GLPIActiveEntities{
    param(
        $Session = (Get-GlpiSession)
    )

    $uri = $glpiUri + "getActiveEntities/"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $r = Invoke-WebRequest -Headers $headers -Method Get -Uri $uri -UseBasicParsing

    $result = ($r | ConvertFrom-Json)
    $result
}

function Set-GLPIActiveEntities{
    param(
        $Session = (Get-GlpiSession),
        $EntitiesID = 0
    )

    $uri = $glpiUri + "changeActiveEntities/"
    $headers = @{
        "Content-Type" = 'application/json'
        "Session-Token" = $Session
    }
    $body = (@{entities_id = $EntitiesID; is_recursive = $true} | ConvertTo-Json)
    $r = Invoke-WebRequest -Headers $headers -Method Post -Uri $uri -UseBasicParsing -Body $body

    $result = ($r | ConvertFrom-Json)
    $result
}

# FIN GLPI