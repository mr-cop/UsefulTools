param (
    [Parameter(Mandatory = $true)]
    [string]$Domain,

    [switch]$MyDebug,

    # Allow users to specify a custom DNS server (default: 8.8.8.8)
    [string]$DnsServer = "8.8.8.8",

    # Maximum recursion depth for expanding SPF records (default: 5)
    [int]$MaxDepth = 5
)

# Global cache for DNS lookups to minimize repeated queries
if (-not $Global:DNSCache) {
    $Global:DNSCache = @{}
}

# Function to retrieve the SPF record for a given domain
function Get-SPFRecord {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Domain,
        [string]$DnsServer
    )

    # Return cached result if available
    if ($Global:DNSCache.ContainsKey($Domain)) {
        if ($MyDebug) {
            Write-Debug "Returning cached SPF record for $Domain"
        }
        return $Global:DNSCache[$Domain]
    }

    if (-not $Domain) {
        Write-Error "Domain parameter is required."
        return $null
    }

    try {
        $records = Resolve-DnsName -Name $Domain -Type TXT -Server $DnsServer -ErrorAction Stop
        $spfRecord = $records.Strings | Where-Object { $_ -match "^v=spf1" }
        if (-not $spfRecord) {
            Write-Warning "SPF record not found for $Domain."
            return $null
        }
        # Cache the SPF record for future use
        $Global:DNSCache[$Domain] = $spfRecord
        return $spfRecord
    } catch {
        Write-Error "Error retrieving SPF record for $Domain : ${_}"
        return $null
    }
}

# Function to expand the SPF record by resolving 'include:' and 'redirect=' mechanisms
function Expand-SPFRecord {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SPFRecord,
        [string]$DnsServer,
        [int]$Depth = 0,
        [int]$MaxDepth
    )

    if (-not $SPFRecord) {
        Write-Error "SPFRecord parameter is required."
        return $null
    }

    if ($Depth -ge $MaxDepth) {
        if ($MyDebug) {
            Write-Debug "Maximum recursion depth reached at level $Depth. Halting further expansion."
        }
        return $SPFRecord
    }

    # Split the SPF record into tokens and process include/redirect mechanisms
    $tokens = $SPFRecord -split "\s+"
    $newTokens = @()

    foreach ($token in $tokens) {
        if ($token -match "^(include:|redirect=)") {
            # Standardize token if it matches a known case, e.g. Outlook
            if ($token -match "protection\.outlook\.com") {
                if ($MyDebug) { Write-Debug "Standardizing Outlook SPF token: $token" }
                $newTokens += "include:spf.protection.outlook.com"
                continue
            } else {
                $domainFromToken = $token -replace "^(include:|redirect=)", ""
                if ($MyDebug) { Write-Debug "Expanding SPF mechanism '$token' for domain '$domainFromToken'" }
                $includedSPF = Get-SPFRecord -Domain $domainFromToken -DnsServer $DnsServer
                if ($includedSPF) {
                    # Remove the leading "v=spf1" from the included record and split into tokens
                    $includedTokens = ($includedSPF -replace "^v=spf1\s+", "") -split "\s+"
                    $newTokens += $includedTokens
                } else {
                    Write-Warning "Included SPF record not found for domain '$domainFromToken'. Removing token '$token'."
                }
            }
        } else {
            # Keep other tokens unchanged
            $newTokens += $token
        }
    }

    # Reconstruct the SPF record string
    $expandedRecord = "v=spf1 " + ($newTokens -join " ")

    # Recursively expand if nested mechanisms remain
    if ($expandedRecord -match "(include:|redirect=)") {
        return Expand-SPFRecord -SPFRecord $expandedRecord -DnsServer $DnsServer -Depth ($Depth + 1) -MaxDepth $MaxDepth
    } else {
        return $expandedRecord
    }
}

# Function to flatten the SPF record by converting dynamic mechanisms (a, mx) into static ip4/ip6 tokens.
# Other mechanisms (exists, ip4, ip6, all) are retained. The "ptr:" mechanism is left unchanged.
function Flatten-SPFRecord {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$SPFRecord,
        # The base domain used for mechanisms (e.g., a, mx) without explicit domains.
        [Parameter(Mandatory = $true)]
        [string]$BaseDomain,
        [string]$DnsServer
    )

    if (-not $SPFRecord) {
        Write-Error "SPFRecord parameter is required."
        return $null
    }

    $finalTokens = @()
    $tokens = $SPFRecord -split "\s+"

    # Loop over each token in the SPF record
    foreach ($token in $tokens) {
        # Skip the version identifier
        if ($token -eq "v=spf1") { continue }

        # Use regex to parse the token optionally capturing qualifier, mechanism, and domain (if specified).
        if ($token -match '^(?<qualifier>[\+\-\?\~]?)(?<mech>a|mx|ip4|ip6|exists|ptr|all)(:(?<domain>.*))?$') {
            $qualifier = $matches['qualifier']
            if ($qualifier -eq "") { $qualifier = "+" }  # Default qualifier is +
            $mechanism = $matches['mech']
            $mechanismDomain = $matches['domain']

            switch ($mechanism) {
                "a" {
                    $target = if ($mechanismDomain -and $mechanismDomain.Trim() -ne "") { $mechanismDomain } else { $BaseDomain }
                    if ($MyDebug) { Write-Debug "Flattening A mechanism for target: $target" }
                    # Resolve A records (IPv4)
                    try {
                        $aRecords = Resolve-DnsName -Name $target -Type A -Server $DnsServer -ErrorAction Stop
                        foreach ($a in $aRecords) {
                            if ($a.IPAddress -match "^[0-9]+\.[0-9]+\.[0-9]+\.[0-9]+$") {
                                $finalTokens += "${qualifier}ip4:$($a.IPAddress)"
                            }
                        }
                    } catch {
                        Write-Warning "Could not resolve A records for $target in mechanism '$token'."
                    }
                    # Resolve AAAA records (IPv6)
                    try {
                        $aaaaRecords = Resolve-DnsName -Name $target -Type AAAA -Server $DnsServer -ErrorAction Stop
                        foreach ($aaaa in $aaaaRecords) {
                            if ($aaaa.IPAddress -match ":") {
                                $finalTokens += "${qualifier}ip6:$($aaaa.IPAddress)"
                            }
                        }
                    } catch {
                        Write-Warning "Could not resolve AAAA records for $target in mechanism '$token'."
                    }
                }
                "mx" {
                    $target = if ($mechanismDomain -and $mechanismDomain.Trim() -ne "") { $mechanismDomain } else { $BaseDomain }
                    if ($MyDebug) { Write-Debug "Flattening MX mechanism for target: $target" }
                    try {
                        $mxRecords = Resolve-DnsName -Name $target -Type MX -Server $DnsServer -ErrorAction Stop
                        foreach ($mx in $mxRecords) {
                            # The Exchange property holds the MX host
                            $mxHost = $mx.Exchange
                            if ($MyDebug) { Write-Debug "Resolving MX host: $mxHost" }
                            # Resolve A records for the MX host
                            try {
                                $mxARecords = Resolve-DnsName -Name $mxHost -Type A -Server $DnsServer -ErrorAction Stop
                                foreach ($mxA in $mxARecords) {
                                    $finalTokens += "${qualifier}ip4:$($mxA.IPAddress)"
                                }
                            } catch {
                                Write-Warning "Could not resolve A records for MX host $mxHost."
                            }
                            # Resolve AAAA records for the MX host
                            try {
                                $mxAAAARecords = Resolve-DnsName -Name $mxHost -Type AAAA -Server $DnsServer -ErrorAction Stop
                                foreach ($mxAAAA in $mxAAAARecords) {
                                    $finalTokens += "${qualifier}ip6:$($mxAAAA.IPAddress)"
                                }
                            } catch {
                                Write-Warning "Could not resolve AAAA records for MX host $mxHost."
                            }
                        }
                    } catch {
                        Write-Warning "Could not resolve MX records for $target in mechanism '$token'."
                    }
                }
                default {
                    # For ip4, ip6, exists, ptr, all leave as-is.
                    $finalTokens += $token
                }
            } # switch
        }
        else {
            # If the token doesn't match our expected pattern, include it unchanged.
            $finalTokens += $token
        }
    }

    # Reconstruct the flattened SPF record
    $flattened = "v=spf1 " + ($finalTokens -join " ")
    return $flattened
}

# --- Main Script Execution ---
try {
    # Retrieve the original SPF record for the provided domain.
    $spfRecord = Get-SPFRecord -Domain $Domain -DnsServer $DnsServer

    if ($spfRecord) {
        Write-Host "Original SPF record for $Domain :" -ForegroundColor Cyan
        Write-Host $spfRecord -ForegroundColor Cyan

        # Expand the SPF record by recursively resolving include/redirect mechanisms.
        $expandedSPF = Expand-SPFRecord -SPFRecord $spfRecord -DnsServer $DnsServer -MaxDepth $MaxDepth
        Write-Host "`nExpanded SPF record:" -ForegroundColor Yellow
        Write-Host $expandedSPF -ForegroundColor Yellow

        # Flatten the SPF record including handling of a:, mx:, and other mechanisms.
        $flattenedSPF = Flatten-SPFRecord -SPFRecord $expandedSPF -BaseDomain $Domain -DnsServer $DnsServer
        Write-Host "`nFlattened SPF record for $Domain :" -ForegroundColor Green
        Write-Host $flattenedSPF -ForegroundColor Green
    } else {
        Write-Warning "No SPF record found for $Domain."
    }
} catch {
    Write-Error "An unexpected error occurred: $_"
}
