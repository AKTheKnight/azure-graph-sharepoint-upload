# Graph Upload Sample – App-Only Plan (Sites.Selected)

## Identities to Create

1. **Local development service principal**
   - Azure AD app registration, e.g., `graph-upload-local`.
   - Authentication: client secret (simple) or certificate (better for rotation).
   - Used in `ClientSecretCredential` locally; store secret in local secret store (not source).

2. **Managed identities (one per environment)**
   - Enable system-assigned MI on each environment resource (App Service, Function, Container App, etc.).
   - Optional: use user-assigned MIs if you need to share identities across multiple resources inside the same environment.
   - Result: three service principals (dev/uat/prod) + one local app registration.

## Permissions Strategy – Microsoft Graph

- Use **Sites.Selected** application permission to limit SharePoint access to the target site(s).
- Both the local app registration and each managed identity must be granted `Sites.Selected` (application permission) in Entra ID **and** receive site-level assignments.

### Steps per identity
1. **Grant Sites.Selected** in Entra ID:
   - API Permissions → Add → Microsoft Graph → Application Permissions → `Sites.Selected`.
   - Grant admin consent.
2. **Assign SharePoint site access** via Microsoft Graph (administrator role required):
   - Use `Sites/{site-id}/permissions` endpoint to grant the identity read/write. Example request body:
     ```json
     {
       "roles": ["write"],
       "grantedToIdentities": [
         {
           "application": {
             "id": "<client-or-mi-object-id>",
             "displayName": "graph-upload-local"
           }
         }
       ]
     }
     ```
   - Repeat for each identity (local SP, dev MI, uat MI, prod MI) and each SharePoint site/library that sample should touch.

## Coding Implications

- Local: `new ClientSecretCredential(tenantId, clientId, clientSecret)`.
- Azure: `new ManagedIdentityCredential()` with optional clientId parameter if using user-assigned MI.
- Scopes remain `https://graph.microsoft.com/.default`; tokens contain Sites.Selected permission + site-specific grants.

## Operational Notes

- Rotations: rotate client secret regularly; managed identities rotate automatically.
- Site changes: when new SharePoint sites are introduced, repeat the `sites/{id}/permissions` assignment for each identity.
- Validation: use `GET /sites/{id}/permissions` to confirm assignments.
