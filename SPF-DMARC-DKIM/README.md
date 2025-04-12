# SPF Record Expander & Flattener

This PowerShell script retrieves the SPF (Sender Policy Framework) record for a given domain, expands any nested `include:` and `redirect=` mechanisms, and then “flattens” the record by converting dynamic mechanisms (such as `a:` and `mx:`) into static IP tokens (`ip4:` and `ip6:`). It retains other mechanisms (such as `exists:`, `ptr:`, and `all`) as necessary.

## Features

- **Retrieve SPF Record:** Resolves DNS TXT records to extract the SPF record.
- **Expand SPF Record:** Recursively expands nested `include:` and `redirect=` mechanisms.
  - *Maximum Recursion:* Controlled by the `-MaxDepth` parameter (default is 5) to prevent infinite loops.
- **Flatten SPF Record:** Converts dynamic mechanisms:
  - **`a:` Mechanism:** Resolves A (IPv4) and AAAA (IPv6) records.
  - **`mx:` Mechanism:** Retrieves MX records, then resolves their A/AAAA records.
  - **Other Mechanisms:** `ip4:`, `ip6:`, `exists:`, `ptr:`, and `all` are retained as is.
- **DNS Caching:** Caches DNS responses to improve performance.
- **Configurable DNS Server:** Use a custom DNS server (default: 8.8.8.8) via the `-DnsServer` parameter.
- **Debug Output:** Enable detailed logging with the `-MyDebug` switch.
- **Base Domain Parameter:** Used in the flattening step for mechanisms lacking an explicit domain.

## Prerequisites

- **PowerShell Version:** PowerShell 5.0 or later (or PowerShell Core).
- **Network Access:** The script requires network access to perform DNS lookups.

## Usage

1. **Run the Script:**

   Open PowerShell and execute the script with the required parameters. For example:

   ```powershell
   .\SPFRecordExpander.ps1 -Domain "example.com" [-MyDebug] [-DnsServer "8.8.8.8"] [-MaxDepth 5]
