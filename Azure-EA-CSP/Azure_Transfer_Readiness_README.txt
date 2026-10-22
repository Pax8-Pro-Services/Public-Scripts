WARNING: Use these tools at your own risk. They are provided "as is" without warranties or guarantees.
Review and validate the scripts and their output before relying on them; you are responsible for resulting decisions and actions.

AZURE TRANSFER READINESS - UNIFIED DISCOVERY v2.2
October 6, 2026

ONE SCRIPT
Azure_Transfer_Readiness_Check.ps1 is the only required script.
Azure_Transfer_Validation_SelfService_Guide_v2.2.pptx is the matching guide.
The full source is also recorded in slide 6 speaker notes for reference.
Use the supplied .ps1 to run it, not text copied from speaker notes.
Azure_Transfer_Readiness_QuickCheck.txt is an optional launcher for the SAME script,
not a separate or reduced inventory implementation.

HOW TO RUN
1. Open https://shell.azure.com and choose PowerShell.
2. Confirm the current tenant and subscription with: az account show
3. Upload Azure_Transfer_Readiness_Check.ps1 and work in that folder.
4. Run:
   & (Join-Path (Get-Location).Path 'Azure_Transfer_Readiness_Check.ps1')
5. Review AzureTransferReadiness.txt and AzureTransferReadiness.json.
   Download them through the Cloud Shell toolbar if required.

OPTIONAL PARAMETERS
-TenantId "<tenant-guid>"
  Must match the current signed-in tenant. Does not switch tenants or sign in.
-SubscriptionIds "<subscription-guid>","<subscription-guid>"
  Narrows resource discovery. Inaccessible or excluded requested IDs remain
  explicitly skipped/unknown; they are not silently removed.
-OutputPath ".\reports\readiness"
  File prefix, absolute or relative. Writes <prefix>.json and <prefix>.txt;
  creates parent folders. Existing reports at the same prefix are replaced.

Local PowerShell 5.1 or PowerShell 7 is also supported with Azure CLI installed
and an existing Azure public-cloud sign-in. No module installation is performed.

WHAT IT DISCOVERS
- Live caller-visible subscriptions in one service tenant; billing account,
  agreement, service/billing tenant evidence and source provenance.
- Reservation/RI orders and children, SKU, quantity, state, expiry and renewal
  where returned. Known parent child references are reconciled against details.
- Savings-plan orders and children, state, term, expiry and renewal where returned.
- Purchasing billing subscription versus benefit applied scope. Shared,
  management-group, unknown and outside-selected scopes are distinguished.
- ARG-indexed Marketplace resource plans and Microsoft.SaaS records, with
  publisher/offer/plan and available state, term and renewal metadata.
- MCA/MPA billing products and possible support products. These records are
  billing-account-wide, not proven selected-subscription purchases.
- INDICATIVE transfer route per subscription (subscriptions[].transferRoute):
  Pay-As-You-Go, Direct EA, Indirect EA, MCA-E (Field-led), MCA online
  (Customer-led) or already CSP, with the pathway to CSP:
    PAYG, Indirect EA, Customer-led MCA: Pathway 1 (offer move to a Field-led
      MCA-E, field supported) then Pathway 2 (partner billing transfer to CSP).
    Direct EA, Field-led MCA-E: Pathway 2 only.
  Offer blockers are flagged from the subscription offer (quotaId / SKU):
  Dev/Test, Visual Studio / MSDN / MPN, Sponsorship, Free / Lightweight Trial,
  student/promotional offers; Azure in Open needs confirmation.
  Conflicting or missing signals are reported as needs-confirmation, never guessed.
- Per-subscription TRANSFER SCOPE (subscriptions[].transferScope): reservations
  and savings plans billed to or applied to each subscription, shared-scope
  benefits to allocate, and Marketplace resources on it. complete=false means
  the lists are lower bounds.
- Subscription state: Disabled/Deleted subscriptions are blocked-by-state;
  Warned/PastDue need confirmation (clear past-due balance first).
- Billing tenant different from service tenant is called out per subscription.
- Large estates: Resource Graph is queried in batches of up to 1,000 subscriptions.

HOW TO READ THE RESULTS
The JSON schema is azure-transfer-readiness/2.2 (2.0 plus transferRoute/quotaId and transferScope).
coverage.<category> contains status, count, observedCount, errors and a note.
  complete-for-visible-scope: all available pages collected for this caller.
  partial: some discovery, detail, reference or scope information is incomplete.
  unavailable: the source could not be collected reliably.
  not-applicable: this API does not support the known agreement/scope.
Incomplete counts are null (shown as unknown in text). observedCount records
the rows actually retrieved. A successful empty list is zero observed, not
proof that no purchases are owned. Per-order coverage includes unresolved
child identifiers if the order lists benefits whose detail is missing.

IMPORTANT LIMITS
- One signed-in service tenant per run; run separately for other tenants.
- Subscription Reader alone does not provide billing or benefit read access.
  Use appropriate read roles on the subscriptions, billing accounts and benefits.
- The billing-products API supports MCA/MPA, not EA/MOSP. Unsupported purchase
  records, usage-billed products and off-platform support require reconciliation.
- ARG is a resource index, not a complete purchase ledger. Metadata can be
  absent or delayed. Do not sum ARG and billing-product rows as unique purchases.
- Caller-visible benefits may relate to other subscriptions. Visibility alone
  does not mean a benefit should be transferred.
- The transfer route is INDICATIVE. It uses offer and billing-account metadata
  (EA enrollment channel / indirect partner, MCA account type), never credit
  card or invoice settings. Confirm the authoritative channel in CST.
- CSP-enabled Marketplace products may transfer with their subscriptions.
  SaaS does not move with its associated Azure subscription: arrange a separate
  transfer and confirm eligibility. Confirm offer/publisher rules, private offers,
  renewal and dependencies, including third-party AI and network-appliance
  licensing where present. Do not infer eligibility from inventory alone.
- Support output is a candidate list, not a definitive support inventory.
  Azure support does not transfer. Confirm contracts and the correct
  renewal/cancellation action manually, including whether all subscriptions move.
- Discovery is not a guarantee of CST equivalence, transfer eligibility,
  destination compatibility, or approval. Confirm the eligible transfer route
  and request grouping with the partner/Microsoft transfer process.
- No live customer pilot was performed during this revision. Before production
  use, compare a known EA/MCA estate with its authoritative purchase records.

SAFETY AND SHARING
No Azure writes, transfers, role changes, sign-in or context changes are performed.
Data requests are allowlisted ARM GETs and Resource Graph POSTs with bounded
retries, pagination and partial-result detection. HTTP redirects are disabled.
An existing CLI sign-in supplies a tenant-checked token held only in memory.
Reports do not intentionally export credentials, tokens, identity email, costs,
payment instruments or raw resource configuration. They DO contain subscription,
resource and billing IDs/names plus selected metadata. Inspect before sharing
through an approved channel; apply appropriate sensitivity labeling to the
presentation before distribution. Reports are never automatically sent.

VERSION NOTE
Use the v2.2 deck and unified script together. Slide 8 covers the indicative transfer route; the v2 deck is superseded. The earlier deck and customer
guide are superseded for this workflow and are not included in the v2 kit.
Original files were preserved in backup\transfer-validation-20260910.

TRANSFER GUIDANCE
https://learn.microsoft.com/en-us/azure/cost-management-billing/manage/mpa-request-ownership
https://learn.microsoft.com/en-us/marketplace/transfer-eligibility
Endpoint references are in the script header.
