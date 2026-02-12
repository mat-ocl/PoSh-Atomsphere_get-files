if($PSVersionTable.PSVersion.Major -lt 3)
{
    Write-Host "Powershell version needs to be a least 3.0!"
    Exit 100
}
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12


# --- Create the needed login string for Api-call headers ---
function BoomiLogin
{
    $password = ConvertTo-SecureString $token -AsPlainText -Force
    $pwdstring = [System.Text.Encoding]::UTF8.GetBytes(($user + ":" + ((New-Object PSCredential "user",$password).GetNetworkCredential().Password)))
    $login = [System.Convert]::ToBase64String($pwdstring)
    return $login
}

# --- Get Boomi Connectors of the Execution ---
function Get-BoomiConnectorsForExecution
{
    # needs to be an array in json
    $executionIDs = @()
    $executionIDs += $executionID
    
    # query expression
    $body = @{"QueryFilter" = 
                @{
                    "expression"=
                    @{"argument" = $executionIDs;"operator" = "EQUALS";"property" = "executionId"}
                }
             }

    # get the connectors
    $executionconnectorquery = "https://api.boomi.com/api/rest/v1/$accountid/ExecutionConnector/query"
    $execConnector = Invoke-Restmethod -Method POST -Uri $executionconnectorquery -Body ($body | ConvertTo-Json -Depth 7) -Headers @{"Accept" = "application/json";"Authorization" = "Basic " + $login} -ContentType "application/json"
    return $execConnector
}

# --- Get the information of the Record from a Connector ---
function Get-BoomiRecords
{
    PARAM( [string]$connectorname ="",
           [string]$recordID ="",
           [string]$executionID = "",
           [string]$accountID = "",
           [string]$login = "" )

    # needs to be an array in json
    $executionIDs = @()
    $executionIDs += $executionID

    $processes = @()
    $processes += $recordID
   
    $connectornames = @()
    $connectornames += $connectorname

    $body = @{"QueryFilter" =
                @{
                   "expression" = @{"nestedExpression" = 
                                    @{"argument" = $executionIDs;"operator" = "EQUALS";"property" = "executionId"},
                                    @{"argument" = $processes;"operator" = "EQUALS";"property" = "executionConnectorId"},
                                    @{"argument" = $connectornames;"operator" = "EQUALS";"property" = "operationName"};"operator" = "and"
                                }
                }
            }

    $GenericConnectorRecordUrl = "https://api.boomi.com/partner/api/rest/v1/$accountid/GenericConnectorRecord/query"
    $execRecord = Invoke-Restmethod -Method POST -Uri $GenericConnectorRecordUrl -Body ($body | ConvertTo-Json -Depth 7) -Headers @{"Accept" = "application/json";"Authorization" = "Basic " + $login} -ContentType "application/json"
    
    $recs = [System.Collections.ArrayList]@()
    $null = $execRecord.result | ForEach-Object {$recs.Add($_)}

    $queryToken = $execRecord.queryToken
    while($null -ne $queryToken)
    {
        $GenericConnectorRecordQueryMoreUrl = "https://api.boomi.com/partner/api/rest/v1/$accountid/GenericConnectorRecord/queryMore"
        $queryMore = Invoke-Restmethod -Method POST -Uri $GenericConnectorRecordQueryMoreUrl -Body $queryToken -Headers @{"Accept" = "application/json";"Authorization" = "Basic " + $login} -ContentType "application/json"
        $null = $queryMore.result | ForEach-Object {$recs.Add($_)}
        $queryToken = $queryMore.queryToken
    }

    return $recs
}

# --- Get the actual documents ---
function Get-BoomiConnectorDocument
{
    PARAM( $rec,
           [string]$fileName = "",
           [string]$Path = "",
           [string]$login = "")

    $myIndex = $rec.incrementalDocumentIndex | ForEach-Object ToString 000
    $myID = $rec.id
    
    # --- Create a document object ---
    $body = @{"genericConnectorRecordId" = $myID }
    Write-Host "Request document from Atomsphere"
    $CreateConnectorDocumentUrl = "https://api.boomi.com/partner/api/rest/v1/$accountid/ConnectorDocument"
    $execCreateConnectorDocument = Invoke-Restmethod -Method POST -Uri $CreateConnectorDocumentUrl -Body ($body | ConvertTo-Json -Depth 7) -Headers @{"Accept" = "application/json";"Authorization" = "Basic " + $login} -ContentType "application/json"

    $fileName = $Path + "\" + $myIndex + "_" + $fileName + ".txt"

    # --- As a reply we get the url where the document can be fetched ---
    Write-Host "Link to document:" $execCreateConnectorDocument.url
    Write-Host "Waiting for document creation.."
    Start-Sleep -Seconds 2
    Write-Host "Attempting a download"
    Invoke-Restmethod -Method GET -Uri $execCreateConnectorDocument.url -Headers @{"Accept" = "application/json";"Authorization" = "Basic " + $login} -ContentType "application/json" -OutFile $fileName

    # --- Query the Url until we get the file ---
    while([String]::IsNullOrWhiteSpace((Get-content $fileName))){
        Write-Host "Retrying.."
        Invoke-Restmethod -Method GET -Uri $execCreateConnectorDocument.url -Headers @{"Accept" = "application/json";"Authorization" = "Basic " + $login} -ContentType "application/json" -OutFile $fileName
        if([String]::IsNullOrWhiteSpace((Get-content $fileName)))
        {
            Write-Host "Waiting.."
            Start-Sleep -Seconds 1
        }
    }
    Write-Host "Saving to disk" $fileName
}

# --- For fixing invalid Paths ---
function FixPath
{
    PARAM([string]$path="")
    return $path.Split([IO.Path]::GetInvalidFileNameChars()) -join '_'
}

# --- Get all Documents for a list of Connectors ---
function Get-AllConnectorDocuments
{
    # Let's make a folder for the executionID
    New-Item -ItemType Directory -Force -Path $executionID | Out-Null
    
    # Package functions for Parallel execution - this is because the Scope is a bit wonky with these
    $GetBoomiRecords = ${function:Get-BoomiRecords}.ToString()
    $GetBoomiConnectorDocument = ${function:Get-BoomiConnectorDocument}.ToString()
    $FixPath = ${function:FixPath}.ToString()

    $connectors.result | Foreach-Object -ThrottleLimit 5 -Parallel { # FIX the []-part
        
        # Unpack functions for parallel use
        ${function:Get-BoomiRecords} = $using:GetBoomiRecords
        $GetBoomiConnectorDocument = $using:GetBoomiConnectorDocument
        ${function:FixPath} = $using:FixPath

        # Unpack variables for parallel use
        $login = $using:login
        $accountID = $using:accountID
        $fileName = FixPath -path $_.executionConnector
        $executionID = $using:executionID

        $records = Get-BoomiRecords -connectorname $_.executionConnector -recordID $_.id -executionID $_.executionId -accountID $accountID -login $login
        Write-Host $records.Count "Records found for " $_.executionConnector

        $records | ForEach-Object -ThrottleLimit 5 -Parallel { 
            ${function:Get-BoomiConnectorDocument} = $using:GetBoomiConnectorDocument
            $accountid = $using:accountID
            
            Get-BoomiConnectorDocument -rec $_ -Path $using:executionID -fileName $using:fileName -login $using:login
        }
    }
}
