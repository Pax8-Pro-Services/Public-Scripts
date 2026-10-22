# WARNING: Use this tool at your own risk. It is provided "as is" without warranties or guarantees.
# Review and validate its output before relying on it; you are responsible for resulting decisions and actions.
<#
.SYNOPSIS
Read-only, caller-visible Azure billing-transfer discovery. NOT an eligibility test.
.DESCRIPTION
Requires an existing Azure CLI sign-in in Azure public cloud. Never signs in, changes
the active account, or transfers anything. One signed-in tenant per run. A CLI token
is acquired for the validated subscription and its tenant is checked. ARM GETs and
Resource Graph POSTs use that token in memory; redirects are disabled. No token is
printed, written to a file, or passed as a process argument. Other tenants are excluded.
Subscription Reader and separate billing/reservation/savings-plan roles may be needed.

Output schema azure-transfer-readiness/2.2: <OutputPath>.json and <OutputPath>.txt,
UTF-8 without BOM. OutputPath is a file PREFIX, absolute or relative to the working
directory (parent directories are created). Existing reports at that prefix are replaced.
Arrays contain OBSERVED metadata only. coverage.<category>.count is null on incomplete
or failed collection; observedCount retains successfully collected rows. Even
complete-for-visible-scope means only the caller-visible records, NOT a tenant census.
Orders retain referencedChildIds from documented parent references, reconciled after
all child pages. unresolvedChildIds identifies references without valid child details;
these IDs are not detail objects or counts. Invalid reference shapes are explicit gaps.
Text reports include order IDs, child coverage, unresolved IDs and per-item metadata gaps.
Review required always: authoritative channel, hidden purchases, and final eligibility.
2.1 adds subscriptions[].transferRoute: an INDICATIVE PAYG / Direct EA / Indirect EA /
MCA / already-CSP route plus offer blockers (Dev/Test, Visual Studio, Sponsorship, Free
Trial) from subscriptionPolicies.quotaId and billing-account enrollmentDetails/accountType.
2.2 adds subscriptions[].transferScope (RIs / savings plans billed or applied per
subscription, shared benefits, Marketplace resources), state-based flags (non-Enabled
subscriptions), MPN / Lightweight Trial / Azure in Open offer handling, tenant-split
summary, Resource Graph batching (<=1,000 subscriptions per request) and progress output.
Reports include IDs/names and selected commercial metadata; inspect before sharing.
No token, identity email, raw resource configuration, payment details or costs exported.

.EXAMPLE
.\Azure_Transfer_Readiness_Check.ps1 -TenantId "<tenant-guid>" -SubscriptionIds "<sub-guid>" -OutputPath ".\reports\readiness"

Endpoint/schema references (GET except Resource Graph POST):
https://learn.microsoft.com/en-us/rest/api/resources/subscriptions/list?view=rest-resources-2022-12-01
https://learn.microsoft.com/en-us/rest/api/billing/billing-property/get?view=rest-billing-2024-04-01
https://learn.microsoft.com/en-us/rest/api/billing/billing-accounts/get?view=rest-billing-2024-04-01
https://learn.microsoft.com/en-us/rest/api/billing/products/list-by-billing-account?view=rest-billing-2024-04-01
https://learn.microsoft.com/en-us/rest/api/reserved-vm-instances/reservation-order/list?view=rest-reserved-vm-instances-2022-11-01
https://learn.microsoft.com/en-us/rest/api/reserved-vm-instances/reservation/list?view=rest-reserved-vm-instances-2022-11-01
https://learn.microsoft.com/en-us/rest/api/billingbenefits/savings-plan-order/list?view=rest-billingbenefits-2022-11-01
https://learn.microsoft.com/en-us/rest/api/billingbenefits/savings-plan/list?view=rest-billingbenefits-2022-11-01
https://learn.microsoft.com/en-us/rest/api/azureresourcegraph/resourcegraph/resources/resources?view=rest-azureresourcegraph-resourcegraph-2022-10-01
https://raw.githubusercontent.com/Azure/azure-rest-api-specs/main/specification/saas/resource-manager/Microsoft.SaaS/Saas/preview/2018-03-01-beta/saas.json
#>
[CmdletBinding()]
param(
    [string[]] $SubscriptionIds,
    [string] $OutputPath = 'AzureTransferReadiness',
    [string] $TenantId
)

$ErrorActionPreference = 'Stop'
$ProgressPreference = 'SilentlyContinue'
$VerbosePreference = 'SilentlyContinue'
$DebugPreference = 'SilentlyContinue'
$script:ArmHost = 'https://management.azure.com'
$script:GraphUrl = "$script:ArmHost/providers/Microsoft.ResourceGraph/resources?api-version=2022-10-01"
$script:Warnings = New-Object System.Collections.ArrayList
$script:Utf8 = New-Object System.Text.UTF8Encoding($false)
$script:Anchor = $null
$script:AuthTenant = $null
$script:AuthToken = $null
$script:TokenAcquired = [DateTime]::MinValue
$script:SelectedIds = @()
$script:UnresolvedIds = @()
$script:GuidPattern = '[0-9a-fA-F]{8}-(?:[0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}'

function Test-Guid($Value) { return ([string]$Value -match "^$script:GuidPattern$") }
function Get-Field($Object, [string]$Name) {
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { return $Object.$Name }
    return $null
}
function Test-Object($Value) { return ($Value -is [pscustomobject]) }
function Text($Value) {
    if ($null -eq $Value) { return $null }
    if ($Value -isnot [string] -and $Value -isnot [ValueType]) { return $null }
    $s = [string]$Value
    $s = $s -replace '(?i)[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}', '[redacted-email]'
    $s = $s -replace '(?i)Bearer\s+\S+', '[redacted-token]'
    $s = $s -replace 'eyJ[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+\.[A-Za-z0-9_-]+', '[redacted-token]'
    return ($s -replace '[\x00-\x1f]', ' ')
}
function Warn([string]$Message) { [void]$script:Warnings.Add($Message) }
function Failure([string]$Kind) { return [pscustomobject]@{ ok=$false; data=$null; error=$Kind } }
function Get-ErrorKind([string]$Message) {
    switch -Regex ($Message) {
        'InvalidAuthenticationToken|ExpiredAuthenticationToken|AADSTS|AuthenticationFailed|Please run.*az login|\b401\b' { return 'auth' }
        'AuthorizationFailed|InsufficientPermission|Forbidden|does not have authorization|\b403\b' { return 'permission' }
        'TooManyRequests|Throttl|ServiceUnavailable|GatewayTimeout|InternalServerError|temporar|timed? ?out|ConnectionError|\b(408|429|500|502|503|504)\b' { return 'transient' }
        'ResourceNotFound|NotFound|Not Found|\b404\b' { return 'notfound' }
        'BadRequest|InvalidResourceType|InvalidApiVersion|NoRegisteredProvider|InvalidQuery|InvalidRequest|\b400\b' { return 'badrequest' }
        default { return 'error' }
    }
}
function Invoke-Cli([string[]]$Arguments) {
    # No raw CLI error text enters either output. Native stderr handling works in PS5/7.
    $old = $ErrorActionPreference
    $ErrorActionPreference = 'Continue'
    $PSNativeCommandUseErrorActionPreference = $false
    try {
        $global:LASTEXITCODE = 0
        $raw = & az @Arguments 2>&1
        $code = $LASTEXITCODE
        $text = (@($raw) -join "`n").Trim()
        if ($code -ne 0) { return (Failure (Get-ErrorKind $text)) }
        if ([string]::IsNullOrWhiteSpace($text)) { return (Failure 'malformed') }
        try {
            $data = ConvertFrom-Json -InputObject $text -ErrorAction Stop
            # Validate standalone JSON first, then preserve arrays via a property:
            # PS7 enumerates top-level []/[one], unlike PS5. No -NoEnumerate in PS5.
            if ($text -match '^\[') {
                $data = (ConvertFrom-Json -InputObject ('{"value":' + $text + '}') -ErrorAction Stop).value
            }
            if ($null -eq $data) { return (Failure 'malformed') }
            return [pscustomobject]@{ ok=$true; data=$data; error=$null }
        } catch { return (Failure 'malformed') }
    } catch { return (Failure (Get-ErrorKind $_.Exception.Message)) }
    finally { $ErrorActionPreference = $old }
}
function Test-ArmUrl([string]$Url) {
    $u = $null
    if (-not [uri]::TryCreate($Url, [UriKind]::Absolute, [ref]$u)) { return $false }
    if ($u.Scheme -cne 'https' -or $u.Host -ine 'management.azure.com' -or
        $u.Port -ne 443 -or $u.UserInfo -or $u.Fragment -or $Url.Contains('\')) { return $false }
    # Restrict both initial and continuation URLs to this inventory's read endpoints.
    $path = $u.AbsolutePath
    return ($path -match '^/subscriptions/?$' -or
        $path -match "^/subscriptions/$script:GuidPattern/providers/Microsoft\.Billing/billingProperty/default$" -or
        $path -match '^/providers/Microsoft\.Billing/billingAccounts/[^/]+(?:/products)?$' -or
        $path -match '^/providers/Microsoft\.Capacity/reservationOrders(?:/[^/]+/reservations)?$' -or
        $path -match '^/providers/Microsoft\.BillingBenefits/savingsPlanOrders(?:/[^/]+/savingsPlans)?$' -or
        $path -eq '/providers/Microsoft.ResourceGraph/resources')
}
function Invoke-Arm([string]$Url, [string]$Method='get', $Body=$null) {
    if (-not (Test-ArmUrl $Url)) { return (Failure 'unsafe-url') }
    if ($Method -ne 'get' -and -not ($Method -eq 'post' -and $Url -eq $script:GraphUrl)) {
        return (Failure 'unsafe-method')
    }
    for ($attempt=0; $attempt -lt 3; $attempt++) {
        # az rest ignores global --subscription for tenant-level authentication.
        # Explicit token acquisition binds get_raw_token to our validated account.
        if (-not $script:AuthToken -or ([DateTime]::UtcNow - $script:TokenAcquired).TotalMinutes -ge 4) {
            $auth = Invoke-Cli @('account','get-access-token','--subscription',$script:Anchor,
                '--resource',"$script:ArmHost/",'--only-show-errors','-o','json')
            if (-not $auth.ok) { $r = Failure $auth.error }
            elseif ((Get-Field $auth.data 'tenant') -ine $script:AuthTenant -or
                (Get-Field $auth.data 'subscription') -ine $script:Anchor -or
                (Get-Field $auth.data 'tokenType') -ine 'Bearer' -or
                (Get-Field $auth.data 'accessToken') -isnot [string] -or
                [string]::IsNullOrWhiteSpace($auth.data.accessToken)) { $r = Failure 'auth' }
            else {
                $script:AuthToken = $auth.data.accessToken
                $script:TokenAcquired = [DateTime]::UtcNow
                $r = $null
            }
            $auth = $null
            if ($null -ne $r) {
                if ($r.error -eq 'transient' -and $attempt -lt 2) { Start-Sleep -Seconds ([int][math]::Pow(2,$attempt)); continue }
                return $r
            }
        }
        $headers = @{ Authorization="Bearer $script:AuthToken" }
        try {
            $request = @{
                Uri=$Url; Method=$Method; Headers=$headers; UseBasicParsing=$true
                MaximumRedirection=0; TimeoutSec=90; ErrorAction='Stop'
            }
            if ($null -ne $Body) {
                $request.ContentType='application/json'
                $request.Body=$script:Utf8.GetBytes((ConvertTo-Json -InputObject $Body -Depth 12 -Compress))
            }
            $response = Invoke-WebRequest @request
            if ([int]$response.StatusCode -ne 200) { $r = Failure (Get-ErrorKind ([string]$response.StatusCode)) }
            else {
                try {
                    $text = $response.Content
                    if ($text -is [byte[]]) { $text=$script:Utf8.GetString($text) }
                    $data = ConvertFrom-Json -InputObject $text -ErrorAction Stop
                    if ($null -eq $data) { $r = Failure 'malformed' }
                    else { $r=[pscustomobject]@{ ok=$true; data=$data; error=$null } }
                } catch { $r = Failure 'malformed' }
            }
        } catch {
            $status = $null
            if ($_.Exception.Response) { $status=[int]$_.Exception.Response.StatusCode }
            $kind = Get-ErrorKind "$status $($_.Exception.Message)"
            if ($status -ge 300 -and $status -lt 400) { $kind='unsafe-redirect' }
            $r = Failure $kind
        } finally {
            $headers.Clear()
            $request=$null; $response=$null
        }
        if ($r.error -eq 'auth' -and $attempt -lt 2) { $script:AuthToken=$null; continue }
        if ($r.ok -or $r.error -ne 'transient' -or $attempt -eq 2) { return $r }
        Start-Sleep -Seconds ([int][math]::Pow(2,$attempt))
    }
}
function New-Coverage([string]$Status, [int]$Observed, [string[]]$Errors=@(), [string]$Note='') {
    return [pscustomobject]@{
        status=$Status
        count=$(if ($Status -eq 'complete-for-visible-scope') { $Observed } else { $null })
        observedCount=$Observed
        errors=@($Errors | Select-Object -Unique)
        note=$Note
    }
}
function Read-List([string]$Url) {
    $rows = New-Object System.Collections.ArrayList
    $seen = @()
    $errors = @()
    $pages = 0
    $firstPath = ([uri]$Url).AbsolutePath
    while ($Url) {
        # nextLink query strings are opaque; compare case-sensitively.
        if ($seen -ccontains $Url -or $pages -ge 10000) { $errors += 'pagination-cycle-or-limit'; break }
        if (-not (Test-ArmUrl $Url) -or ([uri]$Url).AbsolutePath -ine $firstPath) { $errors += 'unsafe-nextlink'; break }
        $seen += $Url
        $r = Invoke-Arm $Url
        if (-not $r.ok) { $errors += $r.error; break }
        # Direct property access preserves empty/singleton JSON array identity.
        $v = $r.data.value
        if (-not (Test-Object $r.data) -or $v -isnot [array]) { $errors += 'malformed'; break }
        $pages++
        foreach ($item in $v) {
            if (-not (Test-Object $item)) { $errors += 'malformed-item'; continue }
            [void]$rows.Add($item)
        }
        $next = Get-Field $r.data 'nextLink'
        if ($null -ne $next -and $next -isnot [string]) { $errors += 'malformed-nextlink'; break }
        $Url = $next
    }
    return [pscustomobject]@{ items=@($rows); pages=$pages; errors=@($errors); complete=($errors.Count -eq 0) }
}
function List-Coverage($Result, [int]$Observed, [string[]]$Extra=@(), [string]$Note='') {
    $e = @(@($Result.errors) + @($Extra) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    $status = if ($e.Count -eq 0) { 'complete-for-visible-scope' } elseif ($Result.pages -gt 0) { 'partial' } else { 'unavailable' }
    return (New-Coverage $status $Observed $e $Note)
}
function Read-Properties([string]$Url) {
    $r = Invoke-Arm $Url
    if ($r.ok -and (-not (Test-Object $r.data) -or -not (Test-Object (Get-Field $r.data 'properties')))) {
        return (Failure 'malformed')
    }
    return $r
}
function Scope-Subscription($Scope) {
    if (Test-Guid $Scope) { return ([string]$Scope).ToLowerInvariant() }
    if ([string]$Scope -match "^/subscriptions/($script:GuidPattern)(?:/|$)") { return $Matches[1].ToLowerInvariant() }
    return $null
}
function Membership($Properties) {
    $billing = Scope-Subscription (Get-Field $Properties 'billingScopeId')
    $scopeType = Text (Get-Field $Properties 'appliedScopeType')
    $asp = Get-Field $Properties 'appliedScopeProperties'
    $applied = @()
    $sid = Scope-Subscription (Get-Field $asp 'subscriptionId')
    if ($sid) { $applied += $sid }
    $rg = Scope-Subscription (Get-Field $asp 'resourceGroupId')
    if ($rg) { $applied += $rg }
    foreach ($scope in @(Get-Field $Properties 'appliedScopes')) {
        $sid = Scope-Subscription $scope
        if ($sid) { $applied += $sid }
    }
    $applied = @($applied | Select-Object -Unique)
    $billMatch = if ($billing -and $script:UnresolvedIds -notcontains $billing) {
        $script:SelectedIds -contains $billing
    } else { $null }
    $appliedMatch = if ($scopeType -eq 'Single' -and $applied.Count) {
        if (@($applied | Where-Object { $script:SelectedIds -contains $_ }).Count) { $true }
        elseif (@($applied | Where-Object { $script:UnresolvedIds -contains $_ }).Count) { $null }
        else { $false }
    } else { $null }
    $relation = if ($billMatch -eq $true -or $appliedMatch -eq $true) { 'related-review-required' }
        elseif ($billMatch -eq $false -and $appliedMatch -eq $false) { 'outside-selected-subscriptions' }
        else { 'unknown-or-shared-review-required' }
    return [pscustomobject]@{
        billingSubscriptionId=$billing; billingSubscriptionSelected=$billMatch
        appliedSubscriptionIds=$applied; appliedSubscriptionSelected=$appliedMatch
        unresolvedRequestedSubscriptionIds=@(@($billing) + $applied | Where-Object { $_ -and $script:UnresolvedIds -contains $_ } | Select-Object -Unique)
        relation=$relation
        note='Billing purchase scope is distinct from benefit application. Shared/management-group scopes require manual allocation; visibility alone does not establish transfer scope.'
    }
}
function Read-Benefits([string]$Provider, [string]$OrderType, [string]$ChildType) {
    $rootPath = "/providers/$Provider/$OrderType"
    $root = Read-List "$script:ArmHost$rootPath`?api-version=2022-11-01"
    $orders = New-Object System.Collections.ArrayList
    $children = New-Object System.Collections.ArrayList
    $extra = @()
    $orderSeen = @{}
    $orderSequence = @()
    $childSeen = @{}
    foreach ($ord in $root.items) {
        $id = $ord.id
        if ($id -isnot [string] -or $id -notmatch "^$([regex]::Escape($rootPath))/[^/?#%\s\\]+$" -or
            -not (Test-Object $ord.properties)) { $extra += 'malformed-order'; continue }
        $id = Text $id
        if ($orderSeen.ContainsKey($id)) {
            $extra += 'duplicate-order'
            # Duplicate pages must not erase additional parent reference evidence.
            $orderSeen[$id].referenceProperties += @($ord.properties)
            continue
        }
        $orderSeen[$id] = [pscustomobject]@{ order=$ord; referenceProperties=@($ord.properties) }
        $orderSequence += $id
    }
    foreach ($id in $orderSequence) {
        $ord = $orderSeen[$id].order
        $p = $ord.properties
        $list = Read-List "$script:ArmHost$id/$ChildType`?api-version=2022-11-01"
        $childErrors = @()
        $referencedChildIds = @()
        foreach ($op in $orderSeen[$id].referenceProperties) {
            if (-not $op.PSObject.Properties[$ChildType]) { continue }
            # RI: reservations[].id objects; SP: savingsPlans[] string IDs (2022-11-01).
            $references = $op.$ChildType
            if ($references -isnot [array]) { $childErrors += 'malformed-child-references'; continue }
            foreach ($reference in $references) {
                $referenceId = $null
                if ($ChildType -eq 'reservations') {
                    if (Test-Object $reference) { $referenceId = $reference.id }
                } elseif ($reference -is [string]) { $referenceId = $reference }
                if ($referenceId -isnot [string] -or
                    $referenceId -notmatch "^$([regex]::Escape($id))/$ChildType/[^/?#%\s\\]+$") {
                    $childErrors += 'malformed-child-reference'
                    continue
                }
                $referenceId = Text $referenceId
                if ($referencedChildIds -notcontains $referenceId) { $referencedChildIds += $referenceId }
            }
        }
        $orderChildren = 0
        foreach ($child in $list.items) {
            $cid = $child.id
            if ($cid -isnot [string] -or $cid -notmatch "^$([regex]::Escape($id))/$ChildType/[^/?#%\s\\]+$" -or
                -not (Test-Object $child.properties)) { $childErrors += 'malformed-child'; continue }
            $cid = Text $cid
            if ($childSeen.ContainsKey($cid)) { $childErrors += 'duplicate-child'; continue }
            $childSeen[$cid] = $true
            $cp = $child.properties
            $scope = Get-Field $cp 'appliedScopeProperties'
            $expiry = Text (Get-Field $cp 'expiryDateTime')
            if (-not $expiry) { $expiry = Text (Get-Field $cp 'expiryDate') }
            $renew = Get-Field $cp 'renew'
            if ($renew -isnot [bool]) { $renew = $null }
            $row = [pscustomobject]@{
                orderId=$id; id=$cid; displayName=(Text (Get-Field $cp 'displayName'))
                sku=(Text (Get-Field (Get-Field $child 'sku') 'name'))
                state=(Text (Get-Field $cp 'provisioningState'))
                displayState=(Text (Get-Field $cp 'displayProvisioningState'))
                term=(Text (Get-Field $cp 'term')); billingPlan=(Text (Get-Field $cp 'billingPlan'))
                expiryDate=$expiry; renew=$renew; benefitStartTime=(Text (Get-Field $cp 'benefitStartTime'))
                billingScopeId=(Text (Get-Field $cp 'billingScopeId'))
                appliedScopeType=(Text (Get-Field $cp 'appliedScopeType'))
                appliedScopes=@(Get-Field $cp 'appliedScopes' | ForEach-Object { Text $_ } | Where-Object { $_ })
                appliedScopeProperties=[pscustomobject]@{
                    subscriptionId=(Text (Get-Field $scope 'subscriptionId'))
                    resourceGroupId=(Text (Get-Field $scope 'resourceGroupId'))
                    managementGroupId=(Text (Get-Field $scope 'managementGroupId'))
                    tenantId=(Text (Get-Field $scope 'tenantId'))
                }
                membership=(Membership $cp)
                metadataGaps=@()
                transferAction='REVIEW required; do not assume benefit transfers with its billing or applied subscription.'
            }
            if ($ChildType -eq 'reservations') {
                $row | Add-Member reservationId $cid
                $row | Add-Member reservationOrderId $id
                $row | Add-Member quantity (Text (Get-Field $cp 'quantity'))
            } else {
                $row | Add-Member savingsPlanId $cid
                $row | Add-Member savingsPlanOrderId $id
            }
            foreach ($field in @('state','expiryDate','renew','billingScopeId','appliedScopeType')) {
                if ($null -eq $row.$field -or [string]$row.$field -eq '') { $row.metadataGaps += "$field not returned; confirm manually" }
            }
            [void]$children.Add($row)
            $orderChildren++
        }
        $unresolvedChildIds = @($referencedChildIds | Where-Object { -not $childSeen.ContainsKey($_) })
        if ($unresolvedChildIds.Count) { $childErrors += 'unresolved-child-references' }
        $cc = List-Coverage $list $orderChildren $childErrors 'Caller-visible child details reconciled against known parent references after all pages. References are not detail counts; missing properties still require manual review.'
        if ($unresolvedChildIds.Count -or @($childErrors | Where-Object { $_ -like 'malformed-child-reference*' }).Count) {
            # Parent evidence remains even if child access fails before its first page.
            $cc.status = 'partial'
        }
        $extra += @($cc.errors)
        $opExpiry = Text (Get-Field $p 'expiryDateTime')
        if (-not $opExpiry) { $opExpiry = Text (Get-Field $p 'expiryDate') }
        $orderRow = [pscustomobject]@{
            id=$id; displayName=(Text (Get-Field $p 'displayName')); state=(Text (Get-Field $p 'provisioningState'))
            term=(Text (Get-Field $p 'term')); billingPlan=(Text (Get-Field $p 'billingPlan')); expiryDate=$opExpiry
            billingScopeId=$(if ($ChildType -eq 'savingsPlans') { Text (Get-Field $p 'billingScopeId') } else { $null })
            renew=$null; renewalSource='child.renew; no documented order-level renew field'
            childrenCoverage=$cc
            referencedChildIds=$referencedChildIds; unresolvedChildIds=$unresolvedChildIds
            metadataGaps=@()
        }
        $orderFields = @('state','term','billingPlan','expiryDate')
        if ($ChildType -eq 'savingsPlans') { $orderFields += 'billingScopeId' }
        foreach ($field in $orderFields) {
            if ($null -eq $orderRow.$field -or [string]$orderRow.$field -eq '') { $orderRow.metadataGaps += "$field not returned; confirm manually" }
        }
        if (@($childErrors | Where-Object { $_ -like 'malformed-child-reference*' }).Count) {
            $orderRow.metadataGaps += "properties.$ChildType contains an invalid reference shape; reconcile parent references manually"
        }
        if ($unresolvedChildIds.Count) {
            $orderRow.metadataGaps += 'Referenced child details not returned; confirm child access and reconcile unresolvedChildIds manually'
        }
        [void]$orders.Add($orderRow)
    }
    return [pscustomobject]@{
        orders=@($orders); items=@($children)
        coverage=(List-Coverage $root $children.Count $extra 'Tenant-level CALLER-VISIBLE benefits, including unrelated/shared scopes. Observed counts are not in-scope counts or a tenant census.')
        orderCoverage=(List-Coverage $root $orders.Count @($extra | Where-Object { $_ -in @('malformed-order','duplicate-order') }) 'Caller-visible orders only; see each childrenCoverage.')
    }
}
function Transfer-Route([string]$QuotaId, [string]$Agreement, $Account, [string]$Sku, [string]$State) {
    # Indicative only: maps observed offer/agreement signals onto the PAYG / EA / Indirect EA /
    # MCA -> CSP routes. Never an eligibility decision; CST/partner confirmation is required.
    $evidence = @()
    $offer = if ($QuotaId) { ($QuotaId -replace '_\d{4}-\d{2}-\d{2}$','') } else { $null }
    if ($QuotaId) { $evidence += "subscriptionPolicies.quotaId=$QuotaId" }
    if ($Agreement) { $evidence += "agreementType=$Agreement" }
    if ($State) { $evidence += "state=$State" }
    $blocker = $null
    $review = $null
    switch -Regex ($offer) {
        '^MSDNDevTest$' { $blocker = 'Dev/Test offer: CSP has no Dev/Test equivalent. Convert to a standard Azure Plan first; Dev/Test pricing is lost.' }
        '^(MSDN|MPN|MSDNPlatforms)$' { $blocker = "Entitlement offer (${offer}: Visual Studio / MSDN / partner benefit): tied to the subscriber or partner benefit and does not transfer. Move resources to a paid subscription." }
        '^Sponsored$' { $blocker = 'Azure Sponsorship: the sponsorship benefit does not exist in CSP.' }
        '^(FreeTrial|LightweightTrial)$' { $blocker = "Trial offer ($offer): must be upgraded/converted before any transfer." }
        '^(AzureForStudents|AzureForStudentsFree|Students|AzurePass|BizSpark|BizSparkPlus)$' { $blocker = "Benefit/promotional offer ($offer): does not transfer to CSP." }
        '^Internal$' { $blocker = 'Microsoft internal subscription: not a customer transfer candidate.' }
        '^AzureInOpen$' { $review = 'Azure in Open (legacy reseller offer): confirm the route with the transfer desk.' }
    }
    if (-not $blocker -and $Sku -match '(?i)dev\s*/?\s*test') {
        $blocker = 'Azure Plan for DevTest: CSP has no Dev/Test equivalent. Convert to a standard Azure Plan first; Dev/Test pricing is lost.'
        $evidence += "sku=$Sku"
    }
    $stateBlocker = $null
    if ($State -and $State -ne 'Enabled') {
        $stateBlocker = if ($State -in @('Warned','PastDue')) {
            "Subscription state $State`: clear any past-due balance and confirm it is active before requesting the transfer."
        } else { "Subscription state $State`: only active subscriptions transfer. Reactivate it or remove it from the request." }
    }
    $eaChannel = Get-Field $Account 'eaChannel'
    $lsp = Get-Field $Account 'indirectPartnerName'
    $acctType = Get-Field $Account 'accountType'
    if ($eaChannel) { $evidence += "enrollmentDetails.channel=$eaChannel" }
    if ($lsp) { $evidence += "enrollmentDetails.indirectRelationshipInfo.displayName=$lsp" }
    if ($acctType) { $evidence += "billingAccount.accountType=$acctType" }
    $twoStep = 'Pathway 1: field-supported offer move to a Field-led MCA-E; then Pathway 2: authorized partner requests the MCA-E -> CSP billing transfer.'
    $oneStep = 'Pathway 2 only: authorized partner requests the billing transfer to CSP (no MCA-E creation).'
    if ($Agreement -eq 'MicrosoftPartnerAgreement' -or $offer -match '^CSP') {
        $route='already-csp'; $label='Already CSP'; $path='Not a PAYG/EA transfer: CSP-to-CSP partner change, handled by the partners.'
    } elseif ($Agreement -eq 'MicrosoftOnlineServicesProgram' -or ($offer -eq 'PayAsYouGo' -and -not $Agreement)) {
        $route='payg'; $label='Pay-As-You-Go (MOSP)'; $path=$twoStep
    } elseif ($Agreement -eq 'EnterpriseAgreement') {
        if ($eaChannel -match '(?i)indirect' -or ($lsp -and -not $eaChannel)) {
            $route='ea-indirect'; $label='Indirect EA'; $path="Treat like PAYG. EA-to-CSP tooling does not support Indirect EA. $twoStep"
        } elseif ($eaChannel -match '(?i)^EaDirect$' -and -not $lsp) {
            $route='ea-direct'; $label='Direct EA'; $path=$oneStep
        } else {
            $route='ea-channel-unknown'; $label='EA (direct/indirect unknown or conflicting)'
            $path='Read the billing profile company name: a licensing solution provider means Indirect EA (two pathways); the customer means Direct EA (Pathway 2 only).'
        }
    } elseif ($Agreement -eq 'MicrosoftCustomerAgreement') {
        if ($acctType -eq 'Enterprise') {
            $route='mca-field-led'; $label='MCA-E (Field-led)'; $path=$oneStep
        } elseif ($acctType -eq 'Individual') {
            $route='mca-customer-led'; $label='MCA online (Customer-led)'; $path=$twoStep
        } else {
            $route='mca-unknown'; $label='MCA (field/customer-led unknown)'; $path='Confirm Field-led vs Customer-led in CST: Field-led = Pathway 2 only; Customer-led = Pathway 1 then 2.'
        }
    } else {
        $route='unknown'; $label='Unknown'; $path='Agreement not visible to the caller. Confirm the billing agreement and channel in CST.'
    }
    $conflict = $route -eq 'already-csp' -and $blocker
    if ($conflict) { $review = "Conflicting evidence: CSP agreement/offer alongside a non-CSP offer ($offer). Confirm the subscription's actual offer in CST." }
    $status = if ($conflict) { 'needs-confirmation' }
        elseif ($blocker) { 'blocked-by-offer' }
        elseif ($stateBlocker -and $State -notin @('Warned','PastDue')) { 'blocked-by-state' }
        elseif ($stateBlocker -or $review -or $route -in @('unknown','ea-channel-unknown','mca-unknown')) { 'needs-confirmation' }
        elseif ($route -eq 'already-csp') { 'not-applicable' } else { 'indicative' }
    return [pscustomobject]@{
        route=$route; label=$label; pathway=$path; offer=$offer; offerBlocker=$blocker
        stateBlocker=$stateBlocker; reviewNote=$review; status=$status
        evidence=@($evidence | ForEach-Object { Text $_ })
        basis='Indicative from caller-visible offer and billing-account metadata. Not eligibility, CST equivalence or approval; confirm the authoritative channel.'
    }
}

if (-not (Get-Command az -ErrorAction SilentlyContinue)) { throw "Azure CLI is required; no sign-in was attempted." }
if ([string]::IsNullOrWhiteSpace($OutputPath)) { throw 'OutputPath must be a nonempty file prefix.' }
$prefix = $ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($OutputPath)
$parent = [IO.Path]::GetDirectoryName($prefix)
[void][IO.Directory]::CreateDirectory($parent)
$ctx = Invoke-Cli @('account','show','--only-show-errors','-o','json')
if (-not $ctx.ok -or -not (Test-Guid (Get-Field $ctx.data 'tenantId')) -or
    -not (Test-Guid (Get-Field $ctx.data 'id'))) { throw 'Valid signed-in subscription context required; no sign-in was attempted.' }
if ((Get-Field $ctx.data 'environmentName') -ne 'AzureCloud') { throw 'Only AzureCloud (management.azure.com) is supported.' }
$signedInTenant = $ctx.data.tenantId
if ($TenantId -and (-not (Test-Guid $TenantId) -or $TenantId -ine $signedInTenant)) {
    throw 'TenantId must match the existing signed-in tenant. Switch context outside this script and rerun.'
}
$TenantId = $signedInTenant
$script:Anchor = $ctx.data.id
$script:AuthTenant = $TenantId
$wanted = @()
foreach ($id in $SubscriptionIds) {
    if ($null -eq $id -or -not (Test-Guid $id.Trim())) { throw 'Each SubscriptionIds entry must be a subscription GUID.' }
    $wanted += $id.Trim().ToLowerInvariant()
}
$wanted = @($wanted | Select-Object -Unique)
Write-Host "Read-only discovery in tenant $TenantId. Manual transfer review is always required."

$cache = Invoke-Cli @('account','list','--all','--only-show-errors','-o','json')
$cacheSubs = @()
if ($cache.ok -and $cache.data -is [array]) { $cacheSubs = @($cache.data) }
else { Warn 'CLI cached subscription listing unavailable/malformed; live ARM listing is authoritative for this run.' }
$excluded = @()
foreach ($s in $cacheSubs) {
    if ((Test-Guid (Get-Field $s 'id')) -and (Get-Field $s 'tenantId') -ine $TenantId) {
        $excluded += [pscustomobject]@{ subscriptionId=$s.id; tenantId=(Text (Get-Field $s 'tenantId')); reason='different-or-unknown-cached-tenant' }
        Warn "Excluded cached subscription $($s.id): different or unknown tenant. Run separately in that tenant."
    }
}
$live = Read-List "$script:ArmHost/subscriptions?api-version=2022-12-01"
$liveById = @{}
$scopeErrors = @()
foreach ($s in $live.items) {
    $sid = Get-Field $s 'subscriptionId'
    if (-not (Test-Guid $sid)) { $scopeErrors += 'malformed-subscription'; continue }
    if ((Get-Field $s 'tenantId') -ine $TenantId) {
        $excluded += [pscustomobject]@{ subscriptionId=$sid; tenantId=(Text (Get-Field $s 'tenantId')); reason='different-or-unknown-live-tenant' }
        Warn "Excluded live subscription ${sid}: service tenant does not match the selected tenant (including delegated access)."
        continue
    }
    if ($liveById.ContainsKey($sid)) { $scopeErrors += 'duplicate-subscription'; continue }
    $liveById[$sid] = $s
}
$targets = if ($wanted.Count) { $wanted } else {
    @(@($liveById.Keys) + @($cacheSubs | Where-Object { $_.tenantId -eq $TenantId -and (Test-Guid $_.id) } | ForEach-Object { $_.id }) | Sort-Object -Unique)
}
$subs = @()
$skipped = @()
foreach ($sid in $targets) {
    if ($liveById.ContainsKey($sid)) { $subs += $liveById[$sid] }
    else {
        $reason = if (@($excluded | Where-Object { $_.subscriptionId -eq $sid }).Count) { 'excluded-tenant' } else { 'not-observed-in-live-tenant-list' }
        $skipped += [pscustomobject]@{ subscriptionId=$sid; status='Skipped'; knowledge='unknown'; reason=$reason }
        $hint = if ($reason -eq 'not-observed-in-live-tenant-list' -and -not $wanted.Count) {
            ' It is in the Azure CLI cache but not visible live (access removed, moved or cancelled). Refresh with az account clear then az login, or rerun with -SubscriptionIds.'
        } else { '' }
        Warn "Subscription $sid was Skipped ($reason); its inventory is unknown.$hint"
    }
}
Write-Host "Validated $($subs.Count) subscription(s); $($skipped.Count) skipped. Reading billing properties..."
$script:SelectedIds = @($subs | ForEach-Object { $_.subscriptionId })
$script:UnresolvedIds = @($skipped | ForEach-Object { $_.subscriptionId })
if ($skipped.Count) { $scopeErrors += 'skipped-subscriptions' }
$scopeCoverage = List-Coverage $live $subs.Count $scopeErrors 'Only selected subscriptions validated in the signed-in tenant. Skipped IDs remain unknown.'

$subReport = @()
$accountMap = @{}
$billingErrors = @()
$payment = @()
foreach ($s in $subs) {
    $sid = $s.subscriptionId
    $url = "$script:ArmHost/subscriptions/$sid/providers/Microsoft.Billing/billingProperty/default?api-version=2024-04-01"
    $bp = Read-Properties $url
    $p = if ($bp.ok) { $bp.data.properties } else { $null }
    if (-not $bp.ok) { $billingErrors += $bp.error }
    $baId = Text (Get-Field $p 'billingAccountId')
    if ($baId -and $baId -notmatch '^/providers/Microsoft\.Billing/billingAccounts/[^/?#%]+$') {
        $baId=$null; $billingErrors += 'malformed-billing-account-id'
    }
    if ($baId -and -not $accountMap.ContainsKey($baId)) {
        $a = Read-Properties "$script:ArmHost$baId`?api-version=2024-04-01"
        $ap = if ($a.ok) { $a.data.properties } else { $null }
        $accountMap[$baId] = [pscustomobject]@{
            billingAccountId=$baId; displayName=(Text (Get-Field $ap 'displayName'))
            agreementType=(Text (Get-Field $ap 'agreementType')); accountStatus=(Text (Get-Field $ap 'accountStatus'))
            accountType=(Text (Get-Field $ap 'accountType')); accountSubType=(Text (Get-Field $ap 'accountSubType'))
            eaChannel=(Text (Get-Field (Get-Field $ap 'enrollmentDetails') 'channel'))
            indirectPartnerName=(Text (Get-Field (Get-Field (Get-Field $ap 'enrollmentDetails') 'indirectRelationshipInfo') 'displayName'))
            primaryBillingTenantId=$(if (Test-Guid (Get-Field $ap 'primaryBillingTenantId')) { $ap.primaryBillingTenantId } else { $null })
            tenantSource='billingAccount.properties.primaryBillingTenantId (2024-04-01); unknown if absent'
            coverage=(New-Coverage $(if ($a.ok) { 'complete-for-visible-scope' } else { 'unavailable' }) $(if ($a.ok) { 1 } else { 0 }) $(if ($a.ok) { @() } else { @($a.error) }))
        }
    }
    $acct = if ($baId) { $accountMap[$baId] } else { $null }
    $billTenant = if ($acct) { $acct.primaryBillingTenantId } else { $null }
    $propTenant = if (Test-Guid (Get-Field $p 'billingTenantId')) { $p.billingTenantId } else { $null }
    $tenantSource = if ($billTenant) { 'billingAccount.properties.primaryBillingTenantId (2024-04-01)' } else { 'unknown' }
    # Both fields are documented in 2024-04-01. Keep provenance; never guess from login.
    if (-not $billTenant -and $propTenant) { $billTenant=$propTenant; $tenantSource='billingProperty.properties.billingTenantId (2024-04-01)' }
    $conflict = $billTenant -and $propTenant -and $billTenant -ine $propTenant
    if ($conflict) { Warn "Billing tenant sources conflict for $sid; resolve manually."; $billTenant=$null; $tenantSource='conflicting-sources-manual-review' }
    $agreement = Text (Get-Field $p 'billingAccountAgreementType')
    $agreementSource = 'billingProperty.properties.billingAccountAgreementType (2024-04-01)'
    if (-not $agreement -and $acct) { $agreement=$acct.agreementType; $agreementSource='billingAccount.properties.agreementType (2024-04-01)' }
    $quotaId = Text (Get-Field (Get-Field $s 'subscriptionPolicies') 'quotaId')
    $route = Transfer-Route $quotaId $agreement $acct ("$(Text (Get-Field $p 'skuDescription')) $(Text (Get-Field $p 'productName'))") (Text (Get-Field $s 'state'))
    $subReport += [pscustomobject]@{
        quotaId=$quotaId; transferRoute=$route
        subscriptionId=$sid; subscriptionName=(Text (Get-Field $s 'displayName')); state=(Text (Get-Field $s 'state'))
        serviceTenantId=$TenantId; serviceTenantSource='ARM subscriptions list tenantId (2022-12-01)'
        billingTenantId=$billTenant; billingTenantSource=$tenantSource
        billingPropertyTenantId=$propTenant; primaryBillingTenantId=$(if ($acct) { $acct.primaryBillingTenantId } else { $null })
        tenantsDiffer=$(if ($billTenant) { $billTenant -ine $TenantId } else { $null })
        billingAccountId=$baId; billingProfileId=(Text (Get-Field $p 'billingProfileId'))
        invoiceSectionId=(Text (Get-Field $p 'invoiceSectionId'))
        agreementType=$agreement; agreementSource=$agreementSource
        productName=(Text (Get-Field $p 'productName')); skuId=(Text (Get-Field $p 'skuId'))
        skuDescription=(Text (Get-Field $p 'skuDescription'))
        subscriptionBillingStatus=(Text (Get-Field $p 'subscriptionBillingStatus'))
        billingPropertiesRead=$bp.ok; billingPropertiesError=$bp.error
    }
    $payment += [pscustomobject]@{
        subscriptionId=$sid; billingProfileId=(Text (Get-Field $p 'billingProfileId'))
        paymentMethodFamily=(Text (Get-Field $p 'billingProfilePaymentMethodFamily'))
        paymentMethodType=(Text (Get-Field $p 'billingProfilePaymentMethodType'))
        metadataSource='billingProperty (2024-04-01); no payment instrument or invoicing preferences collected'
        purchaseMotionSignal='Undetermined'
        basis='Payment method and invoiceEmailOptIn do not establish customer-led/field-led channel. Confirm authoritative channel with Microsoft/partner.'
    }
}
$accounts = @($accountMap.Values | Sort-Object billingAccountId)
Write-Host 'Reading reservations and savings plans (caller-visible, tenant-wide)...'
$ri = Read-Benefits 'Microsoft.Capacity' 'reservationOrders' 'reservations'
$sp = Read-Benefits 'Microsoft.BillingBenefits' 'savingsPlanOrders' 'savingsPlans'

$market = New-Object System.Collections.ArrayList
$graphErrors = @()
$graphPages = 0
$resourceSeen = @{}
$query = @'
Resources
| where isnotempty(plan) or type startswith 'microsoft.saas/'
| project id, subscriptionId, resourceGroup, name, type,
    publisher=tostring(plan.publisher), product=tostring(plan.product), planName=tostring(plan.name),
    saasPublisher=tostring(properties.publisherId), saasOffer=tostring(properties.offerId),
    saasPlan=tostring(properties.skuId), saasStatus=tostring(properties.status),
    saasRenew=tobool(properties.autoRenew), saasTerm=tostring(properties.term.termUnit),
    saasStart=tostring(properties.term.startDate), saasEnd=tostring(properties.term.endDate),
    legacyPublisher=tostring(properties.saasAppPlan.publisher),
    legacyProduct=tostring(properties.saasAppPlan.product), legacyPlan=tostring(properties.saasAppPlan.name)
| order by id asc
'@
if ($subs.Count) {
    # ARG accepts at most 1,000 subscriptions per request; batch and paginate each batch.
    $argBatchSize = 1000
    $batchCount = [int][math]::Ceiling($script:SelectedIds.Count / $argBatchSize)
    Write-Host "Querying Resource Graph for Marketplace / SaaS ($batchCount batch(es))..."
    for ($bi = 0; $bi -lt $script:SelectedIds.Count; $bi += $argBatchSize) {
    $end = $bi + $argBatchSize - 1
    if ($end -ge $script:SelectedIds.Count) { $end = $script:SelectedIds.Count - 1 }
    $batch = @($script:SelectedIds[$bi..$end])
    $token = $null
    $batchRaw = 0
    $graphSeen = @()
    do {
        $options = @{ '$top'=1000; resultFormat='objectArray'; allowPartialScopes=$false }
        if ($token) { $options['$skipToken']=$token }
        $g = Invoke-Arm $script:GraphUrl 'post' @{ subscriptions=$batch; query=$query; options=$options }
        if (-not $g.ok) { $graphErrors += $g.error; break }
        $data = $g.data.data
        $n = Get-Field $g.data 'count'
        $total = Get-Field $g.data 'totalRecords'
        $truncated = Get-Field $g.data 'resultTruncated'
        if ($data -isnot [array] -or $null -eq $n -or $null -eq $total -or
            "$n" -notmatch '^\d+$' -or "$total" -notmatch '^\d+$' -or
            [long]$n -ne $data.Count -or [string]$truncated -notmatch '^(true|false)$') {
            $graphErrors += 'malformed'; break
        }
        $graphPages++
        $batchRaw += [long]$n
        foreach ($r in $data) {
            $id = Text (Get-Field $r 'id')
            $sid = Get-Field $r 'subscriptionId'
            if (-not (Test-Object $r) -or -not $id -or $batch -notcontains $sid -or
                (Scope-Subscription $id) -ne $sid) { $graphErrors += 'malformed-or-out-of-scope-item'; continue }
            if ($resourceSeen.ContainsKey($id)) { $graphErrors += 'duplicate-resource'; continue }
            $resourceSeen[$id]=$true
            $type = Text (Get-Field $r 'type')
            $saas = $type -eq 'microsoft.saas/resources'
            $legacy = $type -eq 'microsoft.saas/applications'
            $renew = Get-Field $r 'saasRenew'
            if (-not $saas -or $renew -isnot [bool]) { $renew=$null }
            [void]$market.Add([pscustomobject]@{
                resourceId=$id; subscriptionId=$sid; name=(Text (Get-Field $r 'name'))
                resourceGroup=(Text (Get-Field $r 'resourceGroup')); type=$type
                publisher=$(if ($saas) { Text $r.saasPublisher } elseif ($legacy) { Text $r.legacyPublisher } else { Text $r.publisher })
                product=$(if ($saas) { Text $r.saasOffer } elseif ($legacy) { Text $r.legacyProduct } else { Text $r.product })
                plan=$(if ($saas) { Text $r.saasPlan } elseif ($legacy) { Text $r.legacyPlan } else { Text $r.planName })
                status=$(if ($saas -or $legacy) { Text $r.saasStatus } else { $null })
                autoRenew=$renew; term=$(if ($saas) { Text $r.saasTerm } else { $null })
                startDate=$(if ($saas) { Text $r.saasStart } else { $null })
                endDate=$(if ($saas) { Text $r.saasEnd } else { $null })
                source='ARG Resources; SaaS fields use Microsoft.SaaS 2018-03-01-beta schema (skuId is plan ID)'
                metadataGap='Missing status/renewal/term fields are unknown. ARG indexing, scrubbing and publisher purchases can leave gaps.'
                transferAction='REVIEW required with publisher/partner; Marketplace is neither automatically moved nor universally blocked. Review third-party AI and network appliance licensing/dependencies where present.'
            })
        }
        $token = Get-Field $g.data '$skipToken'
        if ($null -ne $token -and $token -isnot [string]) { $graphErrors += 'malformed-token'; break }
        if ($token) {
            # Skip tokens are opaque; compare case-sensitively.
            if ($graphSeen -ccontains $token -or $graphPages -ge 10000) { $graphErrors += 'pagination-cycle-or-limit'; break }
            $graphSeen += $token
        } elseif ([string]$truncated -eq 'true' -or $batchRaw -lt [long]$total) {
            # Raw page counts, not accepted rows: rejected rows are reported separately.
            $graphErrors += 'truncated-without-continuation'
        }
    } while ($token)
    }
}
$graphResult = [pscustomobject]@{ errors=$graphErrors; pages=$graphPages }
$graphCoverage = if ($subs.Count) {
    List-Coverage $graphResult $market.Count $(if ($scopeCoverage.status -ne 'complete-for-visible-scope') { @('subscription-scope-incomplete') } else { @() }) 'ARG-indexed resources in selected subscriptions only; not a purchase ledger.'
} else { New-Coverage 'unavailable' 0 @('no-validated-subscriptions') 'No ARG query made; purchase inventory unknown.' }

$products = New-Object System.Collections.ArrayList
Write-Host 'Reading MCA/MPA billing products...'
$productScopes = @()
foreach ($acct in $accounts) {
    $ba = $acct.billingAccountId
    $agreement = $acct.agreementType
    if (-not $agreement) {
        $types = @($subReport | Where-Object { $_.billingAccountId -eq $ba -and $_.agreementType } | Select-Object -ExpandProperty agreementType -Unique)
        if ($types.Count -eq 1) { $agreement=$types[0] }
    }
    if ($agreement -notin @('MicrosoftCustomerAgreement','MicrosoftPartnerAgreement')) {
        $status = if ($agreement -in @('EnterpriseAgreement','MicrosoftOnlineServicesProgram')) { 'not-applicable' } else { 'unavailable' }
        $productScopes += [pscustomobject]@{ billingAccountId=$ba; coverage=(New-Coverage $status 0 @('unsupported-or-unknown-agreement') 'Products API supports MCA/MPA only. This does NOT mean no billed SaaS/support exists.') }
        continue
    }
    $pr = Read-List "$script:ArmHost$ba/products?api-version=2024-04-01"
    $prErrors = @()
    $prCount = 0
    $seenProducts = @{}
    foreach ($item in $pr.items) {
        $id = Text (Get-Field $item 'id')
        $p = Get-Field $item 'properties'
        if (-not $id -or $id -notmatch "^$([regex]::Escape($ba))/products/[^/?#%]+$" -or -not (Test-Object $p)) {
            $prErrors += 'malformed-product'; continue
        }
        if ($seenProducts.ContainsKey($id)) { $prErrors += 'duplicate-product'; continue }
        $seenProducts[$id]=$true
        $name = Text (Get-Field $p 'displayName')
        $pt = Text (Get-Field $p 'productType')
        [void]$products.Add([pscustomobject]@{
            id=$id; billingAccountId=$ba; displayName=$name; productType=$pt
            productTypeId=(Text (Get-Field $p 'productTypeId')); status=(Text (Get-Field $p 'status'))
            autoRenew=(Text (Get-Field $p 'autoRenew')); endDate=(Text (Get-Field $p 'endDate'))
            purchaseDate=(Text (Get-Field $p 'purchaseDate')); skuId=(Text (Get-Field $p 'skuId'))
            skuDescription=(Text (Get-Field $p 'skuDescription')); tenantId=(Text (Get-Field $p 'tenantId'))
            billingProfileId=(Text (Get-Field $p 'billingProfileId')); invoiceSectionId=(Text (Get-Field $p 'invoiceSectionId'))
            subscriptionId=$null; publisher=$null
            membership='billing-account-visible; subscription association unknown'
            supportCandidate=("$pt $name $(Text (Get-Field $p 'productTypeId'))" -match 'support')
            metadataGap='API has no documented subscriptionId/publisher. Null status/autoRenew/endDate require manual confirmation. Support candidate is a label heuristic, not a definitive support inventory.'
            transferAction='REVIEW required; reconcile billed SaaS/support with purchases and contracts. Not deduplicated against ARG.'
        })
        $prCount++
    }
    $productScopes += [pscustomobject]@{ billingAccountId=$ba; coverage=(List-Coverage $pr $prCount $prErrors 'MCA/MPA non-usage-billed products only, across this visible billing account (not limited to selected subscriptions).') }
}
$support = @($products | Where-Object { $_.supportCandidate })
foreach ($item in @($market) + @($products)) {
    $item | Add-Member manualGaps @()
    $fields = @('status','autoRenew','endDate')
    if ($item.PSObject.Properties['resourceId']) { $fields += @('publisher','product','plan','term') }
    foreach ($field in $fields) {
        if ($null -eq $item.$field -or [string]$item.$field -eq '') { $item.manualGaps += "$field not returned; confirm manually" }
    }
    $item.manualGaps += $item.metadataGap
}
$productErrors = @($productScopes | ForEach-Object { $_.coverage.errors })
$missingBilling = @($subReport | Where-Object { -not $_.billingAccountId }).Count -gt 0
$productStatus = 'complete-for-visible-scope'
if (-not $productScopes.Count) { $productStatus='unavailable'; $productErrors += 'no-visible-billing-account' }
elseif (@($productScopes | Where-Object { $_.coverage.status -ne 'complete-for-visible-scope' }).Count) {
    $productStatus = if (@($productScopes | Where-Object { $_.coverage.status -in @('complete-for-visible-scope','partial') }).Count) { 'partial' }
        elseif (@($productScopes | Where-Object { $_.coverage.status -ne 'not-applicable' }).Count -eq 0) { 'not-applicable' } else { 'unavailable' }
}
if ($missingBilling -or $scopeCoverage.status -ne 'complete-for-visible-scope') {
    $productErrors += 'billing-or-subscription-scope-incomplete'
    $productStatus = if ($productScopes.Count) { 'partial' } else { 'unavailable' }
}
$productCoverage = New-Coverage $productStatus $products.Count $productErrors 'Only caller-visible MCA/MPA non-usage-billed products. Hidden purchases, usage-billed products, unsupported agreements and off-platform support remain manual gaps.'
$billingObserved = @($subReport | Where-Object { $_.billingPropertiesRead }).Count
$billingStatus = if ($billingErrors.Count -or $scopeCoverage.status -ne 'complete-for-visible-scope') { if ($billingObserved) { 'partial' } else { 'unavailable' } }
    elseif ($subs.Count) { 'complete-for-visible-scope' } else { 'unavailable' }
$coverage = [ordered]@{
    subscriptions=$scopeCoverage
    billingProperties=(New-Coverage $billingStatus $billingObserved @($billingErrors + $scopeCoverage.errors) 'Successful endpoint reads do not guarantee billing tenant/agreement fields are populated; inspect provenance.')
    billingAccounts=(New-Coverage $(if (-not $accounts.Count) { 'unavailable' } elseif ($missingBilling -or $scopeCoverage.status -ne 'complete-for-visible-scope' -or @($accounts | Where-Object { $_.coverage.status -ne 'complete-for-visible-scope' }).Count) { 'partial' } else { 'complete-for-visible-scope' }) @($accounts | Where-Object { $_.coverage.status -eq 'complete-for-visible-scope' }).Count @($accounts | ForEach-Object { $_.coverage.errors }) 'Referenced visible billing accounts; not all tenant billing accounts.')
    reservationOrders=$ri.orderCoverage; reservations=$ri.coverage
    savingsPlanOrders=$sp.orderCoverage; savingsPlans=$sp.coverage
    marketplace=$graphCoverage; billingProducts=$productCoverage
    supportProducts=(New-Coverage 'partial' $support.Count @('manual-support-reconciliation') 'Support candidates from visible billing products only; no definitive support census.')
    purchaseMotion=(New-Coverage 'unavailable' 0 @('authoritative-channel-required') 'Undetermined; payment/invoice preferences are not channel evidence.')
}
# Per-subscription transfer scope: what the desk must include or action for each subscription.
$benefitCoverageOk = ($ri.coverage.status -eq 'complete-for-visible-scope' -and $sp.coverage.status -eq 'complete-for-visible-scope')
foreach ($s in $subReport) {
    $sid = $s.subscriptionId
    $riBilled = @($ri.items | Where-Object { $_.membership.billingSubscriptionId -eq $sid } | ForEach-Object { $_.id })
    $riApplied = @($ri.items | Where-Object { $_.membership.appliedSubscriptionIds -contains $sid } | ForEach-Object { $_.id })
    $spBilled = @($sp.items | Where-Object { $_.membership.billingSubscriptionId -eq $sid } | ForEach-Object { $_.id })
    $spApplied = @($sp.items | Where-Object { $_.membership.appliedSubscriptionIds -contains $sid } | ForEach-Object { $_.id })
    $shared = @(@($ri.items) + @($sp.items) | Where-Object { $_.appliedScopeType -and $_.appliedScopeType -ne 'Single' } | ForEach-Object { $_.id })
    $mk = @($market | Where-Object { $_.subscriptionId -eq $sid } | ForEach-Object { $_.resourceId })
    $gaps = @()
    if (-not $benefitCoverageOk) { $gaps += "reservations=$($ri.coverage.status); savingsPlans=$($sp.coverage.status)" }
    if ($graphCoverage.status -ne 'complete-for-visible-scope') { $gaps += "marketplace=$($graphCoverage.status)" }
    $s | Add-Member transferScope ([pscustomobject]@{
        reservationsBilledHere=$riBilled; reservationsAppliedHere=$riApplied
        savingsPlansBilledHere=$spBilled; savingsPlansAppliedHere=$spApplied
        sharedScopeBenefitsToReview=$shared; marketplaceResources=$mk
        complete=($gaps.Count -eq 0); coverageGaps=$gaps
        note='Billed here = purchased under this subscription billing scope; applied here = the benefit applies to it. Include both in the transfer review. Shared/management-group benefits need manual allocation. Incomplete coverage means these lists are lower bounds.'
    })
}
foreach ($key in $coverage.Keys) {
    if ($coverage[$key].status -ne 'complete-for-visible-scope') { Warn "${key}: $($coverage[$key].status); observed $($coverage[$key].observedCount). Review coverage errors and manual gaps." }
}
$manual = @(
    'Confirm authoritative customer-led/field-led channel, agreement route, destination partner and final transfer eligibility with Microsoft/partner. This report is not CST equivalence or transfer approval.',
    'Reconcile hidden/other-tenant purchases, permissions, billing profiles, usage-billed/off-platform SaaS and support contracts. Successful empty lists mean zero observed, not none owned.',
    'Review Marketplace publisher/offer eligibility, licensing, private offers and dependencies (including third-party AI and network appliances where present). Do not assume automatic movement or a universal ban.',
    'Review every reservation and savings-plan order/child, purchasing billing scope versus applied scope, shared/management-group allocation, state, expiry and renewal. Accessible tenant-wide benefits are not automatically in scope.',
    'Resolve null/conflicting billing tenants and agreements, skipped subscriptions, missing metadata and incomplete pagination; ARG may be delayed or scrubbed.',
    'Validate renewal/cancellation/transfer handling for SaaS and support manually; do not sum ARG and billing-products rows as unique purchases.',
    'Review required permissions, outstanding commercial obligations, destination offer compatibility, request grouping and any subscription/resource/identity constraints with the authoritative transfer process.'
)
$report = [pscustomobject]@{
    schema='azure-transfer-readiness/2.2'; generatedUtc=[DateTime]::UtcNow.ToString('o'); signedInTenantId=$TenantId
    assessment='MANUAL REVIEW REQUIRED - discovery only; transfer eligibility undetermined'
    privacy='Contains subscription/resource/billing IDs, names and allowlisted metadata. No identity email, tokens, costs or raw resource configuration intentionally exported. Inspect names/metadata before sharing.'
    scope=[pscustomobject]@{ tenantId=$TenantId; tokenSubscriptionId=$script:Anchor; authentication='CLI get-access-token --subscription; tenant verified; in-memory bearer; HTTP redirects disabled'; requestedSubscriptionIds=$wanted; inspectedSubscriptionIds=$script:SelectedIds; skippedSubscriptions=$skipped; excludedTenantSubscriptions=$excluded }
    summary=[pscustomobject]@{
        subscriptionsInspected=$subs.Count; subscriptionsSkipped=$skipped.Count
        reservationCount=$ri.coverage.count; savingsPlanCount=$sp.coverage.count
        marketplaceItemCount=$graphCoverage.count; billingProductCount=$productCoverage.count; supportProductCount=$null
        observedReservationCount=$ri.items.Count; observedSavingsPlanCount=$sp.items.Count
        observedMarketplaceItemCount=$market.Count; observedBillingProductCount=$products.Count
        purchaseMotionSignals=@('Undetermined')
        transferRoutes=@($subReport | ForEach-Object { $_.transferRoute.route } | Sort-Object -Unique)
        offerBlockedSubscriptionIds=@($subReport | Where-Object { $_.transferRoute.offerBlocker } | ForEach-Object { $_.subscriptionId })
        stateFlaggedSubscriptionIds=@($subReport | Where-Object { $_.transferRoute.stateBlocker } | ForEach-Object { $_.subscriptionId })
        needsConfirmationSubscriptionIds=@($subReport | Where-Object { $_.transferRoute.status -eq 'needs-confirmation' } | ForEach-Object { $_.subscriptionId })
        tenantSplitSubscriptionIds=@($subReport | Where-Object { $_.tenantsDiffer -eq $true } | ForEach-Object { $_.subscriptionId })
        countBasis='Counts are caller-visible rows, NOT owned or selected-subscription benefit totals. Null means incomplete/unavailable/not applicable. Observed counts are lower bounds.'
    }
    coverage=[pscustomobject]$coverage
    subscriptions=$subReport; billingAccounts=$accounts; paymentMotion=$payment
    reservationOrders=$ri.orders; reservations=$ri.items; savingsPlanOrders=$sp.orders; savingsPlans=$sp.items
    marketplace=@($market); billingProducts=@($products); billingProductScopes=$productScopes; supportProducts=$support
    manualChecks=$manual; warnings=@($script:Warnings)
}
$jsonPath = "$prefix.json"
$txtPath = "$prefix.txt"
$script:AuthToken=$null
[IO.File]::WriteAllText($jsonPath, (ConvertTo-Json -InputObject $report -Depth 20), $script:Utf8)
$lines = @('AZURE TRANSFER READINESS - READ-ONLY DISCOVERY', $report.assessment,
    "Generated UTC: $($report.generatedUtc)", "Tenant: $TenantId", $report.privacy, '',
    "Subscriptions inspected: $($subs.Count); skipped: $($skipped.Count)", '')
$lines += 'INDICATIVE TRANSFER ROUTE (PAYG / EA / Indirect EA / MCA -> CSP; confirm in CST, not an eligibility decision)'
foreach ($s in $subReport) {
    $r = $s.transferRoute
    $lines += "$($s.subscriptionId) | $($s.subscriptionName) | route=$($r.label) | offer=$($r.offer) | state=$($s.state) | status=$($r.status)"
    if ($r.offerBlocker) { $lines += "  OFFER BLOCKER: $($r.offerBlocker)" }
    if ($r.stateBlocker) { $lines += "  STATE: $($r.stateBlocker)" }
    if ($r.reviewNote) { $lines += "  REVIEW: $($r.reviewNote)" }
    if ($s.tenantsDiffer -eq $true) { $lines += "  TENANT SPLIT: billing tenant $($s.billingTenantId) differs from service tenant $($s.serviceTenantId); confirm which tenant accepts the transfer." }
    $lines += "  PATH: $($r.pathway)"
    if ($r.evidence.Count) { $lines += "  EVIDENCE: $($r.evidence -join '; ')" }
}
$lines += @('', 'TRANSFER SCOPE PER SUBSCRIPTION (include or action in the request; lists are lower bounds when coverage is incomplete)')
foreach ($s in $subReport) {
    $t = $s.transferScope
    $lines += "$($s.subscriptionId) | $($s.subscriptionName) | RIs billed=$($t.reservationsBilledHere.Count) applied=$($t.reservationsAppliedHere.Count) | SPs billed=$($t.savingsPlansBilledHere.Count) applied=$($t.savingsPlansAppliedHere.Count) | Marketplace=$($t.marketplaceResources.Count) | shared benefits to review=$($t.sharedScopeBenefitsToReview.Count) | complete=$($t.complete)"
    foreach ($pair in @(@('RI billed here',$t.reservationsBilledHere),@('RI applied here',$t.reservationsAppliedHere),@('SP billed here',$t.savingsPlansBilledHere),@('SP applied here',$t.savingsPlansAppliedHere),@('Marketplace',$t.marketplaceResources))) {
        foreach ($id in $pair[1]) { $lines += "  $($pair[0]): $id" }
    }
    if ($t.coverageGaps.Count) { $lines += "  INCOMPLETE: $($t.coverageGaps -join '; ')" }
}
$lines += @('', 'COVERAGE (observed is not a tenant census)')
foreach ($key in $coverage.Keys) {
    $c=$coverage[$key]
    $countText=if ($null -eq $c.count) { 'unknown' } else { [string]$c.count }
    $lines += "${key}: $($c.status); count=$countText; observed=$($c.observedCount); errors=$($c.errors -join ', ')"
}
$lines += @('', 'SUBSCRIPTION BILLING / TENANT PROVENANCE')
foreach ($s in $subReport) {
    $lines += "$($s.subscriptionId) | $($s.subscriptionName) | agreement=$($s.agreementType) | billingTenant=$($s.billingTenantId) | source=$($s.billingTenantSource)"
    if (-not $s.billingPropertiesRead) { $lines += "  BILLING DETAILS UNKNOWN: $($s.billingPropertiesError); confirm billing access for this subscription." }
    $gaps = @()
    foreach ($field in @('agreementType','billingTenantId','billingAccountId')) {
        if (-not $s.$field) { $gaps += "$field unknown; confirm manually" }
    }
    if ($gaps.Count) { $lines += "  METADATA GAPS: $($gaps -join '; ')" }
}
foreach ($s in $skipped) { $lines += "SKIPPED: $($s.subscriptionId) | $($s.reason) | inventory unknown" }
$lines += @('', 'BILLING ACCOUNTS / PRODUCT SCOPE GAPS')
foreach ($a in $accounts) {
    if ($a.coverage.status -ne 'complete-for-visible-scope') {
        $lines += "BILLING ACCOUNT: $($a.billingAccountId) | $($a.coverage.status); errors=$($a.coverage.errors -join ', '); confirm billing account access."
    }
}
foreach ($ps in $productScopes) {
    if ($ps.coverage.status -ne 'complete-for-visible-scope') {
        $lines += "PRODUCT SCOPE: $($ps.billingAccountId) | $($ps.coverage.status); observed=$($ps.coverage.observedCount); errors=$($ps.coverage.errors -join ', '); reconcile billed purchases manually."
    }
}
$lines += @('', 'BENEFIT ORDERS (parent visibility does not guarantee child access)')
foreach ($o in @($ri.orders) + @($sp.orders)) {
    $c = $o.childrenCoverage
    $countText = if ($null -eq $c.count) { 'unknown' } else { [string]$c.count }
    $lines += "ORDER: $($o.id) | $($o.displayName) | state=$($o.state) | term=$($o.term) | billingPlan=$($o.billingPlan) | expiry=$($o.expiryDate) | children=$($c.status); count=$countText; observed=$($c.observedCount); errors=$($c.errors -join ', ')"
    if ($o.referencedChildIds.Count) { $lines += "  Referenced child IDs: $($o.referencedChildIds -join ', ')" }
    if ($o.unresolvedChildIds.Count) { $lines += "  Unresolved child IDs (details unknown; confirm child access and reconcile manually): $($o.unresolvedChildIds -join ', ')" }
    if ($o.metadataGaps.Count) { $lines += "  METADATA GAPS: $($o.metadataGaps -join '; ')" }
    if ($c.status -ne 'complete-for-visible-scope') { $lines += '  ACTION: Confirm reservation/savings-plan child read access for this order and reconcile all pages/references with its owner; do not treat missing details as zero purchases.' }
}
$lines += @('', 'OBSERVED BENEFITS (review billing scope separately from applied scope)')
foreach ($b in @($ri.items) + @($sp.items)) {
    $lines += "$($b.id) | $($b.displayName) | state=$($b.state) | expiry=$($b.expiryDate) | renew=$(if ($null -eq $b.renew) { 'unknown' } else { $b.renew }) | billingScope=$($b.billingScopeId) | applied=$($b.appliedScopeType) | $($b.membership.relation)"
    if ($b.metadataGaps.Count) { $lines += "  METADATA GAPS: $($b.metadataGaps -join '; ')" }
}
$lines += @('', 'MARKETPLACE / BILLED PRODUCTS - REVIEW REQUIRED, NOT AUTOMATIC TRANSFER')
foreach ($m in $market) {
    $lines += "$($m.resourceId) | $($m.publisher)/$($m.product)/$($m.plan) | status=$($m.status) | renew=$(if ($null -eq $m.autoRenew) { 'unknown' } else { $m.autoRenew }) | end=$($m.endDate)"
    $lines += "  MANUAL GAPS: $($m.manualGaps -join '; ')"
}
foreach ($p in $products) {
    $lines += "$($p.id) | $($p.displayName) | status=$($p.status) | renew=$(if (-not $p.autoRenew) { 'unknown' } else { $p.autoRenew }) | end=$($p.endDate) | subscription association unknown"
    $lines += "  MANUAL GAPS: $($p.manualGaps -join '; ')"
}
$lines += @('', 'MANUAL CHECKS') + $manual + @('', 'WARNINGS') + @($script:Warnings) + @('', "Full per-order/per-scope details: $jsonPath")
[IO.File]::WriteAllText($txtPath, ($lines -join [Environment]::NewLine), $script:Utf8)
Write-Host "Saved: $jsonPath"
Write-Host "Saved: $txtPath"
$sum = $report.summary
Write-Host ("Routes: " + ((@($subReport | Group-Object { $_.transferRoute.label } | ForEach-Object { "$($_.Name)=$($_.Count)" })) -join ', '))
Write-Host "Offer blockers: $(@($sum.offerBlockedSubscriptionIds).Count); state flags: $(@($sum.stateFlaggedSubscriptionIds).Count); needs confirmation: $(@($sum.needsConfirmationSubscriptionIds).Count); warnings: $($script:Warnings.Count)"
Write-Host $report.assessment
