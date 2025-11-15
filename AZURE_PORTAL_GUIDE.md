# Guide to Configuring SharePoint MCP Application

> **⚠️ Attention:** This guide requires Azure AD administrator permissions. Ensure you have the necessary approvals before proceeding.

## 1. Register the Application in Microsoft Entra (Azure AD)

Access the Azure portal: [https://portal.azure.com](https://portal.azure.com)

### Steps:

1. Navigate to **Microsoft Entra ID** (formerly Azure Active Directory)
2. Go to **App registrations** → **New registration**
3. Enter a descriptive name (e.g., `mcp-sharepoint-app`)
4. Select **Accounts in this organizational directory only**
5. Leave **Redirect URI** empty (not needed for service accounts)
6. Click **Register**

### Save these values:
- **Application (client) ID** → This is your `SHP_ID_APP`
- **Directory (tenant) ID** → This is your `SHP_TENANT_ID`

---

## 2. Create Client Secret

1. In your registered app, go to **Certificates & secrets**
2. Click **+ New client secret**
3. Add a description (e.g., "MCP SharePoint Secret")
4. Choose expiration (recommended: 24 months)
5. Click **Add**
6. **⚠️ IMPORTANT:** Copy the **Value** immediately → This is your `SHP_ID_APP_SECRET`
   - You won't be able to see it again!

---

## 3. Configure API Permissions

1. In your app, go to **API permissions**
2. Click **+ Add a permission**
3. Select **SharePoint**
4. Select **Application permissions**
5. Choose **Sites.Selected** (recommended for security) OR **Sites.ReadWrite.All** (for all sites)
6. Click **Add permissions**
7. **⚠️ CRITICAL:** Click **Grant admin consent for [your organization]**
8. Confirm by clicking **Yes**

### Recommended Permission:
- **Sites.Selected** - Allows access only to specific sites you assign (more secure)

### Alternative Permission:
- **Sites.ReadWrite.All** - Allows access to all SharePoint sites (easier setup, less secure)

---

## 4. Assign Permissions to Specific SharePoint Site

**This step is REQUIRED if you chose Sites.Selected in step 3.**

### Option A: Using SharePoint Admin Center (Easiest)

1. Open in browser:
   ```
   https://[your-tenant]-admin.sharepoint.com/_layouts/15/appinv.aspx
   ```
   Example: `https://sofias219-admin.sharepoint.com/_layouts/15/appinv.aspx`

2. Fill the form:
   - **App Id**: Paste your Application (client) ID
   - Click **Lookup** button
   - **App Domain**: Enter your verified domain (e.g., `Sofias.ai`)
   - **Redirect URI**: `https://[your-tenant].sharepoint.com` (e.g., `https://sofias219.sharepoint.com`)

3. In **Permission Request XML**, paste:
   ```xml
   <AppPermissionRequests AllowAppOnlyPolicy="true">
     <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="Write" />
   </AppPermissionRequests>
   ```

4. Click **Create**
5. Click **Trust It** on the confirmation page

✅ **Done!** Your app now has Write access to that specific site.

### Option B: Using Site-specific URL (Alternative)

If you want to assign to a specific site:

1. Open:
   ```
   https://[your-tenant].sharepoint.com/sites/[site-name]/_layouts/15/appinv.aspx
   ```
   Example: `https://sofias219.sharepoint.com/sites/Clientes/_layouts/15/appinv.aspx`

2. Follow the same steps as Option A

---

## 5. Configure Environment Variables

Create a `.env` file with your saved values:

```env
SHP_ID_APP=your-application-client-id
SHP_ID_APP_SECRET=your-client-secret-value
SHP_TENANT_ID=your-directory-tenant-id
SHP_SITE_URL=https://your-tenant.sharepoint.com/sites/your-site
SHP_DOC_LIBRARY=Shared Documents
```

---

## Permission Scopes Reference

### For Site-specific access (Write):
```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
  <AppPermissionRequest Scope="http://sharepoint/content/sitecollection/web" Right="Write" />
</AppPermissionRequests>
```

### For Site Collection Full Control:
```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
  <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="FullControl" />
</AppPermissionRequests>
```

### For Tenant-wide access:
```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
  <AppPermissionRequest Scope="http://sharepoint/content/tenant" Right="Write" />
</AppPermissionRequests>
```

---

## Troubleshooting

### Error: "Access Denied" (403)
- Verify admin consent was granted in step 3
- Check that the app is assigned to the site in step 4
- Ensure the client secret hasn't expired

### Error: "Invalid Client"
- Verify `SHP_ID_APP` and `SHP_TENANT_ID` are correct
- Check that `SHP_ID_APP_SECRET` was copied correctly

### Error: "Site not found"
- Verify `SHP_SITE_URL` is correct and accessible
- Check that `SHP_DOC_LIBRARY` exists in the site

---

## Security Best Practices

1. ✅ Use **Sites.Selected** instead of Sites.ReadWrite.All
2. ✅ Set client secret expiration and rotate regularly
3. ✅ Grant access only to necessary sites
4. ✅ Store credentials securely (never commit to git)
5. ✅ Use separate apps for dev/test/prod environments
