<#
使用例
$hashTable = @{
    Name     = 'Kevin'
    Language = 'PowerShell'
    State    = 'Texas'
}
Build-QueryString -Params $hashTable
#>
function Build-QueryString {
    Param($Params)

    ([PSCustomObject]$Params | `
     Get-Member -MemberType NoteProperty | `
     Select -ExpandProperty Name | `
     %{ [String]::Concat($_, "=", $Params.$_) }
    ) -join '&'
}

function Invoke-LanscopeAPI {
    Param($BaseUri,$ApiToken,$QueryParams,$Proxy,$Path)

    $headers = @{
        'accept' = 'application/json'
        'Accept-Language' = 'ja-JP'
        'Authorization' = "Bearer ${ApiToken}"
    }

    if($Proxy -eq $null) {
        Write-Host 'match'
    }
    
    [System.UriBuilder]$UriBuilder = [System.UriBuilder]::new($BaseUri)
    $UriBuilder.Path = $Path
    if($QueryParams -ne $null) {
        $UriBuilder.Query = Build-QueryString -Params $QueryParams
    }
    Write-Host $UriBuilder.Uri
    
    Invoke-RestMethod -Uri $UriBuilder.Uri -Method Get -Headers $headers -Proxy $Proxy -ProxyUseDefaultCredentials
}

<#
使用例
$proxy = <Proxy>
$BaseUri = <Base Uri>
$ApiToken = <Api Token>
$HardwareAssetId = <Hardware Asset Id>
$ClientId = <Client Id>

Get-Hardwares -BaseUri $BaseUri -QueryParams @{hardware_asset_type = 1} -Proxy $proxy -ApiToken $ApiToken | ConvertTo-Json | Set-Content -Path .\hardwares.json
Get-Hardware -BaseUri $BaseUri -HardwareAssetId $HardwareAssetId -Proxy $proxy -ApiToken $ApiToken | ConvertTo-Json | Set-Content -Path .\hardware.json
Get-Clients -BaseUri $BaseUri -QueryParams @{mr_type = 1;license_state=1} -Proxy $proxy -ApiToken $ApiToken | ConvertTo-Json | Set-Content -Path .\clients.json
Get-Client -BaseUri $BaseUri -ClientId $ClientId -Proxy $proxy -ApiToken $ApiToken | ConvertTo-Json | Set-Content -Path .\client.json
Get-Groups -BaseUri $BaseUri -QueryParams @{mr_type = 1} -Proxy $proxy -ApiToken $ApiToken | ConvertTo-Json | Set-Content -Path .\groups.json
#>
# ハードウェア資産情報一覧
function Get-Hardwares {
    Param($BaseUri,$ApiToken,$QueryParams,$Proxy)

    Invoke-LanscopeAPI -BaseUri $BaseUri -ApiToken $ApiToken -QueryParams $QueryParams -Proxy $Proxy -Path catapisrv/api/v1/assets/hardwares
}

# ハードウェア資産情報
function Get-Hardware {
    Param($BaseUri,$ApiToken,$HardwareAssetId,$Proxy)

    Invoke-LanscopeAPI -BaseUri $BaseUri -ApiToken $ApiToken -Proxy $Proxy -Path "catapisrv/api/v1/assets/hardwares/${HardwareAssetId}"
}

# クライアント一覧
function Get-Clients {
    Param($BaseUri,$ApiToken,$QueryParams,$Proxy)

    Invoke-LanscopeAPI -BaseUri $BaseUri -ApiToken $ApiToken -QueryParams $QueryParams -Proxy $Proxy -Path catapisrv/api/v1/clients
}

# クライアント情報
function Get-Client {
    Param($BaseUri,$ApiToken,$ClientId,$Proxy)

    Invoke-LanscopeAPI -BaseUri $BaseUri -ApiToken $ApiToken -Proxy $Proxy -Path "catapisrv/api/v1/clients/${ClientId}"
}

# グループ一覧
function Get-Groups {
    Param($BaseUri,$ApiToken,$QueryParams,$Proxy)

    Invoke-LanscopeAPI -BaseUri $BaseUri -ApiToken $ApiToken -QueryParams $QueryParams -Proxy $Proxy -Path catapisrv/api/v1/groups
}
