# Pay-as-you-go features via Microsoft Marketplace (SPFx + XSProject API)

> **Status: implemented and verified (2026-10-14), revision 2.** Backend: 300/300 tests, build 0 errors. SPFx: 500/500 tests, no new lint warnings. All changes are uncommitted working-tree edits; there are **no git commits in either repo**. Progress is tracked in "Implementation progress" below.
>
> **What changed in revision 2:** a read-only review of both repos (findings in the "Review log" at the end) turned up:
> - a gate that would lock out every non-SPFx user;
> - a unique index that breaks the second "enable";
> - Tender endpoints that are not behind `[Authorize]`;
> - a risky startup migration;
> - missed feature endpoints;
> - several activation-flow races.
>
> Every one is addressed below. All decisions below were confirmed by the user on 2026-10-13.

## Implementation progress
Each task has an implementer and a separate reviewer. A task is ticked only once its review is clean. The two lanes run side by side, one task at a time in each.

**Backend** (`D:\Projects\TaskXS\BMDeskV2`)
- [x] **B1** Master-DB entities, master migrations, guarded startup runner, `Marketplace` config section, grandfathering SQL. Review clean after one fix round, which corrected the config keys.
- [x] **B2** Feature codes and the Marketplace fulfillment client. Review clean.
- [x] **B3** `FeatureSubscriptionService`, repositories, DI, the test project and the service tests. Review clean after two fix rounds (a webhook conflict bug; a NOT NULL column that a first-time enable left empty).
- [x] **B4** Identity claims (`xs_account_id`, `xs_entra_tid`) and `CurrentAccountId`. Review clean.
- [x] **B5** `SubscriptionController` and `MarketplaceController` (landing page and webhook). Review clean.
- [x] **B6** The `RequireFeature` gate and where it is applied. Review clean after one fix round (supplier actions exempted; a second ED Control config endpoint gated).

**SPFx** (`D:\Projects\SPFx\SPFx`)
- [x] **S1** `FeatureSubscription` model (codes, token reader, activated mapper) and `canManageFeatureSubscriptions`. Review clean.
- [x] **S2** `SubscriptionService`: `returnUrl` on enable, `activateMarketplacePurchase`, cache seeding, test stubs. Review clean.
- [x] **S3** `FeatureSubscriptionSwitch` and `SubscriptionPrompt` purchase flow. Review clean.
- [x] **S4** Subscription page, routing and menu, `EdControl` removal, Teams detection. Review clean.
- [x] **S5** CLAUDE.md "Paid features" section. Review clean.

- [x] **Final review across both repos:** verdict "ready to merge with fixes". The two sides agree end to end, and the change is safe to ship with `EnforceFeatureGates=false`.
- [x] **Final fix wave**, re-reviewed clean:
  - send the Marketplace activation `quantity` only when Microsoft supplied one;
  - stop gating two lookups that ordinary pages use (`Building/GetProjectIsEnableTender` and `TenderTemplate/GetTenderTemplateLookupItems`);
  - restore your CLAUDE.md sentence;
  - SPFx sends `returnUrl` only when an offer link is configured;
  - the switch uses the two-argument `.then`.

### Before anyone sets `EnforceFeatureGates=true` (checklist)
- [ ] Close the security prerequisites listed above: Tender, Setting and Customer endpoints that don't require sign-in; the SQL string concatenation; the connection string carried in the JWT; plaintext secrets.
- [ ] **Old mobile sessions.** `RefreshToken` renews tokens without limit and never adds `xs_account_id` to tokens issued before this change. Force a re-login, or add the claim during refresh. Otherwise those users get 403 on gated endpoints.
- [ ] Run `Grant_Complimentary_Features.sql` for every account that already uses a paid feature (the MVC app included).
- [ ] Ops check: `SELECT COUNT(*) FROM Account WHERE IsDeleted=0 AND ConnectionString IS NULL`. Such accounts share the default database, and so share account attribution.
- [ ] Decide on the four anonymous supplier Tender actions, which are exempt from the gate, as part of the `[Authorize]` prerequisite work.
- [ ] Deploy the API before the SPFx package; an older API answers 404 and every switch shows "Unable to verify".

**Rulings made during implementation** (Microsoft Learn, *Implementing a webhook on the SaaS service*):
- **Webhook token checks.** `aud` must equal `Marketplace:ClientId`, `tid` must equal `Marketplace:TenantId`, and `appid` or `azp` must equal `20e940b3-4c77-4b0b-9a53-9e16a1b010a7`. No separate publisher keys are needed.
- **Auto-activation.** Marketplace may activate a subscription at purchase.
  - Activation skips Microsoft's `activate` call when the subscription is already `Subscribed`.
  - The webhook also acknowledges `Subscribe` and `ChangeQuantity`.
- **Duplicate tenant migration.** Version `20251117051200` is **not** renumbered. It is out of scope, because changing it would affect live tenant databases.

## Context
Tenants should be able to buy four add-on features (ED Control, KYP, Minutes of Meetings, Tender) through Microsoft Marketplace. Once a domain has paid, every user in that domain can use the feature. An admin turns a feature on from a new **Subscription** page in the SPFx web part. If the feature isn't paid for, they are sent to Marketplace, then come back to SPFx, and the API records the purchase against their domain (`Account`).

**Where things stand:**
- **SPFx** (`feature/ImplementSubsscription` @635620b) already has most of the client:
  - `FeatureSubscription.ts` (model), `SubscriptionService`, `FeatureSubscriptionSwitch`, `FeatureSubscriptionGuard`, `SubscriptionPrompt` and `SubscriptionStatus`;
  - an `/ed-control` page holding only the switch.
- **Backend** (`D:\Projects\TaskXS\BMDeskV2`, branch `feature/1177_SubscriptionSPFx` @b6854ae, clean): net10.0, EF Core 10.0.7. It has SPFx login but no subscription code.
  - `feature/1154_SPFxSubscription` (196cb12) has a controller matching the SPFx contract, but it stores data in the tenant DB and has no Marketplace code.
  - `feature/1176_SubscriptionCore` (e843b46) is a large plan/usage system that conflicts with it.
- **Plan:** build on 1177. Reuse the *shape* of 196cb12's controller and model, and 1176's `FluentMigration.cs` namespace filter and test csproj. **Do not merge either branch.**

## Decisions
| # | Decision | Value |
|---|---|---|
| D1 | Features | Exactly `ED_CONTROL, KYP, MINUTES, TENDER` (drop BIM and REVIEW). |
| D2 | Who can manage (enable, disable, buy) | **Admin (1), SuperAdmin (2), ClientAdmin (5)**. See note below. |
| D3 | Marketplace setup | One SaaS offer, with one plan per feature. |
| D4 | Activation flow | API landing page → SPFx → API verifies with Microsoft. |
| D5 | Existing MVC tenants | **Grandfathered** with `Complimentary` rows, written by an ops SQL script before any gate is switched on. |
| D6 | Master DB migrations | Run at **API startup, guarded**: a SQL app lock, try/log, and the API still starts if it fails. |
| D7 | Gate refusal | **HTTP 403** with a `ResponseDetail` body. Safe only because of D5. |
| D8 | API gating | Built, but `EnforceFeatureGates=false` by default. |
| D9 | Teams hosts | Read-only. Buying is offered only in SharePoint, because the return URL can't round-trip into Teams. |

**Note on D2.** This deliberately differs from `canAddEditProject`, which grants MaintenanceManager because every SPFx user who joins an *existing* domain is hard-coded to that role (`AuthenticateController.cs:709`). Buying commits the domain to a charge, so it stays with administrators. No domain is left without an administrator:
- when SPFxLogin auto-creates a domain, its first user is seeded as Admin (`CustomerService.cs:159`, `@UserRole=1`);
- on an existing domain, an MVC administrator promotes users (`UsersController.cs:305`).

Non-managers see read-only status and "Ask an administrator to subscribe." Client (3, the McD manager) is excluded. Alternative if rejected: add MaintenanceManager, as 196cb12 did.

## Prerequisites in other work (not part of this plan, but they block D8 = true)
Pre-existing security gaps. A subscription gate on top of an open endpoint can be bypassed, so these must be fixed before `EnforceFeatureGates` is ever true:
- `TenderController` and `TenderQuestionsController` have `[Authorize]` commented out (`TenderController.cs:18`, `TenderQuestionsController.cs:15`), so every action is anonymous.
- `SettingController` has no class-level `[Authorize]`, and `GetEDControlsConfiguration` returns ED Control client credentials to anonymous callers.
- `CustomerController` has its `[Authorize]` commented out (`UpdateSubscriptionId`, `DeleteCustomer`).
- `JWTAuthenticationFilter.cs:47` builds SQL by string-concatenating a request header.
- The JWT carries the tenant connection string, including the `sa` password, as a claim (`AuthenticateController.cs:625`).
- Plaintext secrets are committed in `appsettings.json`.
- Two tenant migrations share version `20251117051200`:
  - `Migration_Add_IsDocumentPreview_ColumnIn_User_Entity_20251117051200.cs`
  - `Migration_Add_Record_AppSetting_20251117051200.cs`
  - FluentMigrator throws on the duplicate, so the tenant runner is already broken. Renumber one before touching any runner.

## Activation flow
1. **The admin switches feature X on in Subscription** (SharePoint host only, D9). SPFx then:
   1. **inside the click**, opens a blank tab: `const tab = window.open('', '_blank')`. A tab opened after an `await` is popup-blocked.
   2. sends `POST api/Subscription/features/X/enable` with body `{ returnUrl }`. The return URL is `window.location.href` without its fragment; query strings such as the workbench's `?debug=…` are kept.
   3. **on success**, if `needsMarketplaceCheckout`: sets `tab.opener = null; tab.location.href = marketplaceOfferUrl`. **On refusal**: `tab.close()` and shows the error.
   - This order means a refused enable never leaves the user on a Marketplace page with no checkout row behind it.
2. **The API saves the switch and a pending checkout.**
   - It sets `IsEnabled=true` and upserts a **pending checkout**: AccountId, FeatureCode, ReturnUrl, EntraTenantId, CreatedBy, ExpiresAt = now + 24h. There is one live row per (AccountId, FeatureCode); a newer enable replaces it.
   - It validates `ReturnUrl` (this prevents open redirects):
     - scheme is `https`;
     - host is `{sp}.sharepoint.com` or `{sp}-my.sharepoint.com`, where `{sp}` = `Account.Domain` with `DefaultDomain` (`.ifour-consultancy.net`) stripped off the end;
     - no userinfo, and no fragment.
   - Multi-geo and sovereign hosts are out of scope and are refused with a clear message.
3. **The customer pays.** Marketplace sends them to the single Partner Center landing URL, `GET api/Marketplace/landing?token=…` (anonymous).
4. **The landing endpoint routes the purchase.**
   - It resolves the token with Microsoft (read-only) to get the beneficiary tenant, the plan, and from the plan the feature.
   - It finds the newest unexpired pending checkout matching (EntraTenantId, FeatureCode) and answers 302 to `{returnUrl}#/subscription?subscription=returned&marketplaceToken={Uri.EscapeDataString(token)}`.
   - Otherwise it answers a small HTML page:
     - when there is no match (a purchase made on AppSource, an expired checkout, or the wrong plan): "We couldn't link this purchase automatically. Open XSProject → Subscription and switch the feature on, then contact support with reference {subscriptionId}". Nothing is written.
     - when the token is missing, empty or unresolvable: the same page with no reference. **Not** the global `ValidateApiInputFilter` JSON 400, so the action takes an optional parameter and validates it itself.
   - Marketplace does not charge until `activate` is called, so an unlinked purchase costs the customer nothing.
5. **SPFx consumes the token.**
   - The Subscription page sees `marketplaceToken` (read without `+`→space decoding, since Marketplace tokens are base64-like) and calls `POST api/Subscription/marketplace/activate { token }` **once per token**. A module-level `Set` of consumed tokens survives remounts.
   - It then removes the token with `navigate('/subscription?subscription=returned', { replace: true })`.
   - The token is never logged. Never set `print.includeRequest` on this call.
6. **The API activates**, then answers with the feature's subscription:
   1. resolves the token again;
   2. checks that the plan maps to a known feature;
   3. checks that **the caller's AccountId** owns an unexpired checkout for that feature, and that its EntraTenantId equals the token's beneficiary tenant. While SPFx sign-in is single-tenant, the AccountId is the real binding;
   4. **is idempotent**: if this account already has Status=Subscribed with the same MarketplaceSubscriptionId, it answers success without calling Microsoft again. This covers "Change site" re-running activation with a spent token;
   5. calls Marketplace `activate` (planId, quantity);
   6. upserts `AccountFeatureSubscription` (Status=Subscribed, Source=Marketplace, MarketplaceSubscriptionId, PlanId, BeneficiaryTenantId) for the caller's AccountId, so the subscription applies to the whole domain;
   7. deletes the checkout;
   8. invalidates the gate cache for (AccountId, Feature).
7. **SPFx refreshes the switches.**
   - It **seeds** its cache with the activation answer, as `setFeatureEnabled` already does at `SubscriptionService.ts:169`; clearing alone could let an in-flight read cache stale data.
   - It re-keys the switches so they render the new state. Today they keep "being processed", because `isReturningFromMarketplace` stays true after `navigate(replace)` and `componentDidUpdate` never re-reads.
   - A feature is usable only when the server says `isActive`. A client claim can never make it active.
8. **Later changes come from Marketplace.** `POST api/Marketplace/webhook` handles Unsubscribe, Suspend, Reinstate, ChangePlan and Renew.
   - It validates the Entra bearer token: issuer is the publisher tenant, audience is the publisher app id.
   - It **re-reads the subscription from the SaaS API** before writing anything, so a forged payload changes nothing.
   - It finds the row by MarketplaceSubscriptionId, so it never touches Complimentary rows.
   - On a ChangePlan to another feature's plan, it moves the row's FeatureCode only if the account has no row for the target feature; otherwise it logs the conflict and leaves both rows alone.
   - It answers 200 quickly; Marketplace retries anything else.

**Known gap:** if the buyer closes the return tab before the site picker, nothing is activated, and Marketplace doesn't charge. Re-enabling the switch creates a new checkout.

## Backend changes (paths under `D:\Projects\TaskXS\BMDeskV2`)

### Data: master DB, keyed by `Account.Id` (`long`)
Webhook calls carry no tenant JWT, and the static `DbContextFactory.conString` makes switching tenant DBs unsafe, so this data lives in the master DB. `BMDeskMasterDBContext` doesn't use that static.

- `Database/BMDesk.Domain/AccountFeatureSubscription.cs`:
  - `long Id`, `long AccountId`, FeatureCode `nvarchar(32)`, IsEnabled, Status `nvarchar(32)` (Subscribed, PendingFulfillmentStart, Suspended, Unsubscribed or null), Source `nvarchar(16)` (`Marketplace` or `Complimentary`), PlanId, MarketplaceSubscriptionId (`uniqueidentifier`, nullable), BeneficiaryTenantId, Quantity, ModifiedBy, Created, Modified;
  - unique index on (AccountId, FeatureCode);
  - **filtered** unique index on MarketplaceSubscriptionId `WHERE MarketplaceSubscriptionId IS NOT NULL`, since SQL Server allows only one NULL in a plain unique index.
- `Database/BMDesk.Domain/MarketplaceCheckout.cs`: `long Id`, `long AccountId`, FeatureCode, ReturnUrl `nvarchar(2048)`, EntraTenantId, CreatedBy, ExpiresAt. Unique (AccountId, FeatureCode).
- **No navigation properties** on either entity. The master context uses lazy-loading proxies, and a non-virtual navigation breaks the master model, and with it every login.
- **No new columns on `Account`.** `JWTAuthenticationFilter.cs:47-50` reads `SELECT *` and takes column `ItemArray[4]` by position.
- Use `long` keys throughout. `BaseRepository.Get(int)` calls `Find(long)` on the master path, so an `int` key throws; use queries rather than `Get(int)`.
- Add both `DbSet`s to `BMDeskMasterDBContext`.

### Master migrations (D6)
- Add `BMDesk.Domain.Migration/MasterMigrations/Migration_Create_FeatureSubscription_Tables_20261013xxxxxx.cs`, written with FluentMigrator, including the filtered index (`.Execute.Sql` for the `WHERE`).
- Change `FluentMigration.cs`, porting 1176's version (e843b46 `FluentMigration.cs:39-43`):
  - add `InitializeMaster(conn)`;
  - filter each runner by namespace (`TypeFilterOptions`): the tenant runner uses `BMDesk.Domain.Migration.Migrations`, and the master runner uses `BMDesk.Domain.Migration.MasterMigrations`;
  - all 22 tenant migrations are already in `…Migrations`, so the filter drops none;
  - check that `FluentMigratorRunScript` (new tenant provisioning, `CustomerService.cs:247`) cannot pick up `MasterMigrations`.
- Call it from `Startup.Configure` with `BMDeskMasterConnection`:
  - wrap it in `sp_getapplock 'XS_MasterMigrations'`, so IIS overlapped recycles don't race;
  - wrap it in try/catch with logging, so the API starts even if this fails;
  - then the subscription endpoints answer an Error envelope "Subscriptions are temporarily unavailable" while the tables are missing, and the gate fails **open** while `EnforceFeatureGates=false`.
- The first run creates FluentMigrator's `VersionInfo` table in master. The master login needs DDL rights for that first run; otherwise run the migration once by hand.

### Feature codes
`Business/BMDesk.Business/Common/FeatureSubscriptionCodes.cs`: `ED_CONTROL, KYP, MINUTES, TENDER`, plus `IsKnown` (case-insensitive) and `Normalize`.

### Marketplace client: `Business/BMDesk.Services/Marketplace/`
- `MarketplaceOptions`, bound through `IOptions<MarketplaceOptions>` from the `Marketplace` section. **Not `ConfigHelper`**, which reads `appsettings.json` directly and would miss environment variables. Fields: TenantId, ClientId, ClientSecret, OfferId, `PlanFeatures` (planId → feature code), `EnforceFeatureGates` (default false), `GateCacheSeconds` (default 60), `SupportContact`.
- `IMarketplaceFulfillmentClient` / `MarketplaceFulfillmentClient`:
  - a typed `HttpClient` (`AddHttpClient` is already registered at `Startup.cs:87`);
  - a client-credentials token for resource `20e940b3-4c77-4b0b-9a53-9e16a1b010a7`, through MSAL `ConfidentialClientApplication` (`Microsoft.Identity.Web` 4.8.0 is already referenced), which caches it until expiry;
  - `ResolveAsync(token)`, `ActivateAsync(id, planId, qty)` and `GetSubscriptionAsync(id)` against `https://marketplaceapi.microsoft.com/api/saas/subscriptions`, api-version `2018-08-31`;
  - passes a `CancellationToken` through;
  - sends the token in the `x-ms-marketplace-token` header (never in a logged URL), and never logs the token.
- No new NuGet packages: JwtBearer 10.0.7 brings `Microsoft.IdentityModel.Protocols.OpenIdConnect` for webhook validation.

### Service
`Interfaces/Services/IFeatureSubscriptionService.cs` and `Business/BMDesk.Services/FeatureSubscriptionService.cs`:
- Methods:
  - `GetAll(accountId)`
  - `Get(accountId, code)`
  - `SetEnabled(accountId, code, enabled, userId, returnUrl, entraTenantId, accountDomain)`
  - `ActivatePurchase(accountId, token, userId)`
  - `ApplyWebhook(evt)`
  - `IsActive(accountId, code)`
- It returns `FeatureSubscriptionModel` (`Feature, IsActive, IsEnabled, SubscriptionStatus, PlanId, PlanName, RequiresSubscription`), the shape SPFx already parses, with `Feature` always uppercase.
- **isActive** = `IsEnabled && ((Source == Marketplace && Status == Subscribed) || Source == Complimentary)`.
- **Status is written only from Microsoft's answer.** `SetEnabled` never changes Status or Source.
- Complimentary rows answer `SubscriptionStatus = "Subscribed"`, so SPFx shows them as active and offers no checkout.
- `IsActive` goes through `IMemoryCache` for `GateCacheSeconds`, keyed by (AccountId, code), and is invalidated on enable, disable, activate and webhook. Without it the gate would hit the master DB on every request.
- Repositories: `AccountFeatureSubscriptionRepository` and `MarketplaceCheckoutRepository` on `bMDeskMasterDB`, registered in `UnitOfWork`. Register the services, options and typed `HttpClient` in `Startup.RegisterDI`.

### Identity (claims)
- **Append** two custom-named claims at the **end** of the claim list, so existing positional reads (`BaseApiController.cs:34-36` `identity[1]`/`[2]`, `JWTAuthenticationFilter.cs:36` `getClaims[4]`) don't move:
  - `xs_account_id` (the `long` Account.Id), in **every** token-minting path:
    - `Login` (`:151-160`)
    - `LoginWithEmail` (`:617-626`), used by SAML `LoginWithAzure` and `SPFxLogin`
    - `AzureAppLoginWithEmail` (`:528-537`)
    - `VerifyCode` (`:1508-1520`)
    - `RefreshToken` (`:276`), which carries it forward from the old token
    - Missing a path means a 403 for those users once gates are on.
  - `xs_entra_tid`, only when SPFxLogin knows it (`ValidateSPFxTokenAsync`'s principal, `:1179`), carried through a new `LoginViewModel.EntraTenantId`.
- **Do not use `tid`**: JwtBearer's default inbound mapping rewrites it. **Do not reuse any claim type already in the list**: duplicate types serialize into one JSON array at the first occurrence, which shifts the positional indexes.
- `BaseApiController` gets `CurrentAccountId` (a `long?`), read **by claim type**.
- Tokens issued before this change lack the claim; they live 3h with zero clock skew. Until they expire:
  - the subscription endpoints answer an Error envelope "Please sign in again";
  - the gate treats a missing claim as **not active** when enforced.

### Controllers
- `BMDesk.API/Controllers/SubscriptionController.cs`:
  - `[Authorize] api/Subscription` (no clash with the legacy `api/Subscriptions`), `: BaseApiController`, `base(new HttpContextAccessor())`, returning the `ResponseDetail` envelope;
  - endpoints: `GET features`, `GET features/{code}`, `POST features/{code}/enable` (body `{returnUrl}`, required), `POST features/{code}/disable`, `POST marketplace/activate` (body `{token}`);
  - enable, disable and activate require `userRoleId ∈ {1, 2, 5}` (D2); otherwise an Error envelope "Only an administrator can manage subscriptions";
  - an unknown code gets an Error envelope.
  - This controller still runs `JWTAuthenticationFilter`, which sets the static tenant connection string. That's harmless for master-only work; note it, don't fix it here.
- `BMDesk.API/Controllers/MarketplaceController.cs`:
  - a `ControllerBase`, *not* `BaseApiController`, so `JWTAuthenticationFilter` doesn't run, with `[AllowAnonymous]`;
  - `GET landing?token=` (optional parameter; HTML fallback, never JSON) and `POST webhook`;
  - webhook validation uses `JwtSecurityTokenHandler` with the Entra OpenID configuration for `Marketplace:TenantId` (`ConfigurationManager<OpenIdConnectConfiguration>`, which caches the signing keys).
  - The webhook's bearer token will also fail the default JwtBearer scheme. That's harmless log noise; tag it with `[Authorize(AuthenticationSchemes = "none")]`, or simply leave it.

### Gate
`BMDesk.API/Filters/RequireFeatureAttribute.cs`, an `IAsyncActionFilter` via `TypeFilter`:
- When `EnforceFeatureGates` is true and the caller's account isn't `IsActive` for the feature, it answers **403** with a `ResponseDetail` body "This feature requires an active subscription" (D7).
- When the switch is off, or the master tables are missing, it does nothing.
- Apply it as follows:
  - **MINUTES:**
    - `MOMController`, `MOMTemplateController`, `MOMTemplateTopicController`
    - `Setting/GetMOMConfiguration`
    - `Document/AddMOMDocument`, `Document/GetMOMDocument`
  - **TENDER:**
    - `TenderController`, `TenderQuestionsController` (after the prerequisite restores `[Authorize]`)
    - `TenderDocumentController`, `TenderTemplateController`, `TenderSubFolderDocumentController`
    - `Building/GetProjectIsEnableTender`
    - **Not** `TenderQuoteController`, or the anonymous `AuthenticateController` tender endpoints (`:1312-1437`), which are supplier-facing.
    - The per-project `IsEnableTender` flag is unchanged; the gate is account-level on top of it.
  - **ED_CONTROL:**
    - `EDControlAuditController`
    - `Building/GetEDContorlsTickets`, `Building/GetProjectLinkWithEDControl`
    - `Document/GetDocumentDetailsForEdControls`, `Document/GenerateFileUrlForEdControls`, `Document/UpdateEDControlDetails`
    - `Setting/GetEDControlsConfiguration` (after the prerequisite)
    - **Not** `Setting/auditReportCallback` or `Setting/callbackHandler`, which are ED's own callbacks.
  - **KYP:** `Building/GetKYPPlannings`, `Setting/GetKYPProjectConfiguration`.
- **Out of scope:** Hangfire jobs (`Startup.cs:313-330`) keep running for unsubscribed tenants.

### Grandfathering (D5)
- Ship `Database/Scripts/Grant_Complimentary_Features.sql`, a parameterised template that inserts `Source=Complimentary, IsEnabled=1, Status=NULL` rows for (AccountId, FeatureCode). Ops run it for every existing account that already uses a feature, **before** any environment sets `EnforceFeatureGates=true`.
- The API never writes `Complimentary`, so no client path can grant one.
- If a complimentary account later buys, activation converts the row to `Source=Marketplace`.
- This keeps the MVC app working when gates go on. MVC's `DoActionForGetAsync` returns null on a non-success response, and `BuildingController.cs:224-225` calls `.Any()` on that null; a 403 there is a NullReferenceException.

### Config
- Add a `Marketplace` section to `appsettings.json` with empty ids, empty `PlanFeatures` and `EnforceFeatureGates: false`.
- The secret goes in user-secrets for Development only (`UserSecretsId` exists; `Program.cs:14` loads it only in Development) and in an **environment variable** (`Marketplace__ClientSecret`) on IIS. Never put it in the file.

### Tests
Add a new xUnit project, `Tests/BMDesk.Services.Tests`. Recreate it from 1176's csproj (xunit 2.9.3, Test.Sdk 17.14.1, `Microsoft.EntityFrameworkCore.Sqlite` 10.0.7) and add it to `BMDesk/BMDesk.sln` as `..\Tests\…`. Today only stray bin/obj folders sit there.

Use **Sqlite in-memory**, not EF InMemory, which doesn't enforce the unique indexes these tests exist to catch. `UnitOfWork` also needs a `CoreDemoDBContext`, so give it a Sqlite one too. Use a fake `IMarketplaceFulfillmentClient`. Cover:
- **Enable and checkout:**
  - enable doesn't activate;
  - a second account can enable before buying (the filtered index);
  - a bad `returnUrl` host or scheme is refused;
  - an enable replaces the previous checkout.
- **Activation:**
  - a tenant mismatch is refused;
  - an unknown plan is refused;
  - a checkout belonging to another account is refused;
  - an expired checkout is refused;
  - a repeat activation with the same subscription id succeeds without calling Microsoft.
- **isActive:**
  - only `subscribed && enabled`, or `complimentary && enabled`;
  - `SetEnabled` never changes Status.
- **Webhook:**
  - re-reads before writing;
  - ignores Complimentary rows;
  - a ChangePlan collision leaves both rows alone.
- **Claims:** every minting path carries `xs_account_id`, and the positional `[1]`, `[2]` and `[4]` reads are unchanged. This needs a light controller test, or extract the claim-list builder into one method that every path calls.

## SPFx changes (`D:\Projects\SPFx\SPFx\src`)

**`models/FeatureSubscription.ts`**
- `FEATURE_CODES = ['ED_CONTROL','KYP','MINUTES','TENDER']`, named ED Control, KYP, Minutes of Meetings and Tender.
- Add `MARKETPLACE_TOKEN_PARAMETER` and `readMarketplaceToken(search)`, next to `isMarketplaceReturn`. Use `decodeURIComponent` **without** `safeDecode`'s `+`→space.
- Add `toActivatedFeatureSubscription(dto)`: uppercase `dto.feature`, require `isFeatureCode`, then delegate to `toFeatureSubscription`. Return `undefined` otherwise.
- Update the "read-only" docblock (`:9-11`): the client now forwards a Marketplace token, but still never decides activation.
- Update the `#/ed-control` example (`:235`) to `#/subscription`.
- Tests: fix the `'BIM'` `FeatureCode` literal at `FeatureSubscription.test.ts:113`; add cases for `readMarketplaceToken` (`+` preserved, absent → undefined) and `toActivatedFeatureSubscription` (lower case accepted, unknown rejected).

**`services/SubscriptionService.ts`**
- `setFeatureEnabled(feature, isEnabled, returnUrl?)` posts `{ returnUrl }` on enable (today it posts `{}`, typed `Record<string, never>`).
- Add `activateMarketplacePurchase(token)`:
  - posts to `marketplace/activate`, maps the answer with `toActivatedFeatureSubscription`, and gives an `ApiError` on `undefined`;
  - **seeds** the cache for that feature with the answer.
  - Never pass `print.includeRequest`, because the body carries the token.
- Update the docblock (`:27-31`).
- Tests (real `ApiService` over a fake transport, as today):
  - fix the `'BIM'`/`'REVIEW'` literals at `:167,175,183`;
  - add activate success, refusal envelope, unknown feature, and cache seeded.

**`ISubscriptionService` stubs**
Add `activateMarketplacePurchase` to every stub, or the build fails:
- `XsProject.test.tsx:280-285`
- `FeatureSubscriptionGuard.test.tsx:~73`
- `FeatureSubscriptionSwitch.test.tsx:~84`

**`models/Permissions.ts`**
- `canManageFeatureSubscriptions` = Admin, SuperAdmin or ClientAdmin (D2); it no longer delegates to `canAddEditProject`.
- Rewrite its docblock with the D2 note: why this one *doesn't* grant MaintenanceManager, and how a domain gets an Admin.
- **Add** tests (none exist today) for every `UserRole` value (1–10) and for an undefined role.

**`FeatureSubscriptionSwitch.tsx`**
- Use the blank-tab, enable, then navigate-or-close sequence from Activation flow step 1, and pass `returnUrl`.
- Add a prop `isPurchaseAvailable` (false in Teams, D9): when false, enabling a feature that needs checkout shows "Open XSProject in SharePoint to subscribe" instead.
- Keep the `visibilitychange` re-check.
- Update the test assertions at `:174,211,233`, plus new tests for the refused enable (tab closed) and for Teams.

**`SubscriptionPrompt.tsx`**
- The Subscribe link (`:105-118`) must go through the same enable-then-open path, so a checkout row exists, rather than opening Marketplace directly.
- Hide it without `canManage`. Today the switch renders it regardless (`FeatureSubscriptionSwitch.tsx:165`).

**New page `components/Subscription/`**
Files: `Subscription.tsx`, `ISubscriptionProps.ts`, `ISubscriptionState.ts`, `Subscription.module.scss`, `Subscription.test.tsx`.
- **Props:**
  - `subscriptionService`, `canManage`, `isPurchaseAvailable`, `links`
  - `marketplaceToken?`, `isReturningFromMarketplace`
  - `onTokenConsumed`
- **Rendering:**
  - a title and one `FeatureSubscriptionSwitch` per `FEATURE_CODES` entry (data-driven);
  - the switches are keyed by an `activationGeneration` counter, so they re-mount and re-read after activation.
  - Non-managers see the switches disabled, plus "Ask an administrator to subscribe".
- **On mount with a token not in the module-level consumed set:**
  - add it to the set;
  - show "Activating your purchase…";
  - call `activateMarketplacePurchase`;
  - show success ("{Feature} is now active for your organisation") or the API's message;
  - bump `activationGeneration`;
  - call `onTokenConsumed`.
- **Root style** follows the layout rules: `box-sizing: border-box; width: 100%`.
- **Tests:**
  - activation runs once across a remount;
  - success re-keys the switches;
  - a failure shows the message and activates nothing;
  - non-managers are disabled;
  - in Teams, no purchase is offered.

**Routing and menu**
- Add `ROUTE_PATHS.subscription = '/subscription'`.
- In `AppShell.tsx`, add a `<Route>` passing:
  - `marketplaceToken={readMarketplaceToken(location.search)}`;
  - `onTokenConsumed` = `useCallback(() => navigate(ROUTE_PATHS.subscription + '?subscription=returned', { replace: true }))`.
  - AppShell already has `navigate`, `location` and `canManageFeatures` (`AppShell.tsx:46,75-76`), so no extra route adapter is needed.
- Add a `MENU_ITEMS` entry "Subscription".
- **Remove `/ed-control`:**
  - `components/EdControl/*`
  - `Navigation.ts:28`
  - `Menu.tsx:26-28`
  - `AppShell.tsx:12,27,152-163`
  - the comment at `FeatureSubscriptionGuard.tsx:15-16`
  - `XsProject.test.tsx:711-730`: rewrite it against `#/subscription`.

**Teams detection (D9)**
- `XsProjectWebPart` passes `isPurchaseAvailable = !this.context.sdks.microsoftTeams` down through `XsProject` to `AppShell`.
- In Teams, `location.href` is a `_layouts/15/teamshostedapp.aspx` page that won't work outside Teams.

**`CLAUDE.md` "Paid features" section**
- Update the codes, the Subscription page (replacing `EdControl`), the activation flow, the token consumption, D2's manage roles and the Teams rule.
- Remove "this client only reads it" and "no setter"; replace them with "never decides activation".
- Keep the user's existing uncommitted CLAUDE.md edits.

## Prerequisites the user must provide (not code)
- **A Partner Center SaaS offer with four plans.** Set:
  - Landing page URL: `https://<api-host>/api/Marketplace/landing`;
  - Webhook: `https://<api-host>/api/Marketplace/webhook`;
  - the publisher Entra tenant id and app id.
- **An Entra app registration** in the publisher tenant, with a client secret (into an environment variable), and the plan ids for `Marketplace:PlanFeatures`.
- **In SPFx `environment.ts`**, set `subscription.marketplaceOfferUrl` and `manageSubscriptionUrl` per environment. Today they are demo links in local and empty elsewhere.
- **Certification check:** Marketplace certification expects the landing page to identify the buyer with Entra SSO. Here, identification happens after the 302, through the SharePoint sign-in. Confirm with Partner Center that this is acceptable before submitting.
- **Single tenant:** SPFxLogin enforces one `SPFx:TenantId` (`AuthenticateController.cs:1116,1179-1182`), so only one customer Entra tenant can sign in until that becomes multi-tenant. Out of scope; until then, the tenant check in activation step 3 discriminates nothing, and the AccountId binding carries the security.
- **Before switching on gates,** run the grandfathering script (D5) and close the security prerequisites above.

## Verification
- **SPFx:**
  - `heft test` (build, lint and Jest).
  - Focused pass: `heft test --test-path-pattern "FeatureSubscription|SubscriptionService|Subscription/|Permissions|XsProject"`.
  - Run `heft clean` first, since `EdControl` tests are deleted.
- **Backend:**
  - `dotnet build BMDesk/BMDesk.sln`
  - `dotnet test Tests/BMDesk.Services.Tests`
- **Manual (local):**
  1. Start the API. Check the log for master migrations applied (or a logged failure with the API still up), and check that `AccountFeatureSubscription`, `MarketplaceCheckout` and `VersionInfo` exist in master.
  2. Run `npm run start` and sign in as an Admin. Open **Subscription** and check the four switches.
  3. Switch KYP on:
     - a tab opens on the demo Marketplace link;
     - the switch shows "complete your subscription";
     - one `MarketplaceCheckout` row exists, with the workbench URL (query string included) as ReturnUrl.
  4. Switch KYP on for a second account. Its checkout and subscription rows are created too, which shows the filtered index allows it.
  5. Visit `#/subscription?subscription=returned&marketplaceToken=a+b%2Fc`:
     - the API receives `a+b/c` (log the token length, not the value);
     - resolving it fails, and the UI shows the failure;
     - nothing becomes active;
     - the token is gone from the URL;
     - "Change site" and back does not re-send it.
  6. Sign in as MaintenanceManager: the switches are disabled and Subscribe is hidden. Calling `enable` directly returns the refusal envelope.
  7. Use an old token (from before the change, still under 3h old): the subscription endpoints say "Please sign in again".
  8. With `EnforceFeatureGates=true`:
     - a MOM endpoint for an account without a row → 403;
     - after inserting a Complimentary row → 200 (within `GateCacheSeconds`);
     - with the flag false, behaviour is unchanged.
  9. MVC regression: with gates on and a Complimentary row, the MVC project page still loads tender and MOM data.
- **End to end:** a full purchase needs the real Partner Center offer (a preview/test plan in Marketplace) and a deployed API, because the landing URL is fixed.

## Review log (2026-10-13, revision 2)
Each fix below came out of the review:
- **Gate lockout:** `xs_account_id` was only in `LoginWithEmail`, so with gates on every MVC, mobile and supplier user got a 403. Now added in all five minting paths.
- **Unique index:** a plain unique index on the nullable MarketplaceSubscriptionId broke the second enable. Now filtered.
- **Tender gate targets:** "The `[Authorize]` Tender* controllers" missed `TenderController` and `TenderQuestionsController` (their `[Authorize]` is commented out). Now listed explicitly, behind a prerequisite.
- **Startup migration:** "Startup migration, as 1176 did" was wrong (1176 used an admin endpoint). Now a startup run that is locked and fail-safe.
- **Missed feature endpoints:** the gate list missed ED Control, MOM and Tender endpoints in `Building`, `Document` and `Setting`. Now added.
- **ReturnUrl prefix rule:** it ignored `Account.Domain`'s `DefaultDomain` suffix. Now stripped explicitly.
- **Claim names:** `tid` would be remapped, and a reused claim type would shift the positional reads. Now custom names, appended.
- **Key types:** int vs long, and lazy-load proxies. Now `long` keys and no navigation properties.
- **Popup and checkout order:** Marketplace opened before enable was confirmed. Now a blank tab, which is closed on refusal. The Subscribe link bypassed the checkout and `canManage`; now routed through enable.
- **Stale switches** after activation, and **repeat activation** on remount or "Change site": now the cache is seeded, the switches re-keyed, a consumed-token set kept, and activate is idempotent.
- **Token decoding:** `+` was decoded as a space in both the API and SPFx. Now `EscapeDataString`, with no `safeDecode`.
- **Teams:** the return URL can't round-trip. Now read-only there (D9).
- **Test database:** EF InMemory can't test the unique indexes. Now Sqlite.
- **Existing MVC tenants** would lose their features, or crash, once gates went on. Now grandfathered (D5).
