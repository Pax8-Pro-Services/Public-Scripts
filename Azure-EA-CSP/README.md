# Azure EA / CSP Transfer Readiness

> [!WARNING]
> These scripts are provided by Microsoft, but neither Microsoft nor Pax8 provides support for them. Use them at your own risk; they are provided "as is" without warranties or guarantees. Review and validate the scripts and their output before relying on them; you are responsible for decisions and actions taken based on their results. The reports may contain sensitive subscription, resource, and billing information, so inspect them before sharing.

This directory contains a read-only Azure discovery script to help prepare for a subscription billing transfer to a Cloud Solution Provider (CSP). It inventories caller-visible subscriptions and transfer-related billing and benefit information, then produces reports for manual review.

**This is discovery, not an eligibility test or transfer approval.** Confirm the authoritative transfer route, purchase scope, and eligibility with Microsoft and the partner before acting.

## Contents

- `Azure_Transfer_Readiness_Check.ps1`: the full discovery tool. This is the only required script.
- `Azure_Transfer_Readiness_QuickCheck.txt`: optional PowerShell launcher for the same script; it does not perform a smaller or separate inventory.
- `Azure_Transfer_Readiness_README.txt`: detailed feature, coverage, safety, and interpretation notes.
- `Azure_Transfer_Validation_SelfService_Guide_v2.2.pptx`: companion self-service guide for the v2.2 workflow.

Use the `.ps1` file directly. Do not copy script source out of presentation speaker notes.

## Requirements

- Azure CLI installed and already signed in to the Azure public cloud.
- PowerShell 5.1 or PowerShell 7.
- A current Azure CLI context in the tenant to inspect. The script does not sign in, switch tenants, or change the active account.
- Appropriate read permissions for subscriptions and, where applicable, billing accounts, reservations, and savings plans. Subscription Reader alone may not grant access to billing or benefits.

The script performs no Azure writes, transfers, role changes, or sign-in. Its Azure Resource Manager calls are read-only; Resource Graph uses POST for queries.

## Run

In Azure Cloud Shell, select PowerShell, confirm the signed-in context, and upload the script:

```powershell
az account show
& (Join-Path (Get-Location).Path 'Azure_Transfer_Readiness_Check.ps1')
```

Alternatively, from the directory containing the script:

```powershell
.\Azure_Transfer_Readiness_Check.ps1
```

To use the optional quick-check launcher, place both `.ps1` and `.txt` files in the same current directory, then paste the contents of `Azure_Transfer_Readiness_QuickCheck.txt` into PowerShell. It writes reports with the default `AzureTransferReadiness` prefix.

### Optional parameters

```powershell
.\Azure_Transfer_Readiness_Check.ps1 `
  -TenantId '<tenant-guid>' `
  -SubscriptionIds '<subscription-guid>','<subscription-guid>' `
  -OutputPath '.\reports\readiness'
```

- `-TenantId`: validates the existing signed-in tenant; it does not switch tenants or sign in.
- `-SubscriptionIds`: limits resource discovery to the requested subscriptions. Requested subscriptions that cannot be inspected remain identified as skipped or unknown.
- `-OutputPath`: output file prefix, absolute or relative to the working directory. The script creates parent folders and replaces existing reports with the same prefix.

All parameters are optional. `-OutputPath` defaults to `AzureTransferReadiness`.

## Reports

The script writes `<prefix>.json` and `<prefix>.txt` in UTF-8 without a byte-order mark. Review coverage and manual checks in both reports; incomplete or unavailable coverage means the observed inventory may be only a lower bound. A successful empty result does not prove that no purchases exist.

The reports may contain subscription, resource, and billing IDs and names, as well as selected commercial metadata. They do not intentionally export credentials, tokens, identity email, costs, payment instruments, or raw resource configuration. Inspect the reports before sharing them and use an approved channel.

## Important limits

- A run covers one signed-in tenant and only information visible to the caller.
- Reservation and savings-plan visibility, Marketplace indexing, billing-product API scope, permissions, and metadata can all be incomplete.
- Billing-product discovery is limited to supported MCA/MPA non-usage-billed products; it is not a complete purchase or support ledger.
- Marketplace and SaaS transfer rules vary. SaaS does not move with its associated Azure subscription; confirm publisher eligibility, licensing, renewals, and dependencies separately.
- The reported transfer route is indicative. Confirm the authoritative channel, CSP pathway, request grouping, destination compatibility, and final eligibility through the Microsoft/partner transfer process.

See [Azure_Transfer_Readiness_README.txt](Azure_Transfer_Readiness_README.txt) for detailed discovery coverage, report schema, safety notes, and transfer guidance. The script header also lists its endpoint references.