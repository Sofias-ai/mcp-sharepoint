
# Guide to Registering an Application in Azure Portal and Configuring SharePoint Permissions

> **⚠️ Attention:** Many of the steps in this guide may require the permissions of your organization’s owner or an Azure AD administrator. Please follow the instructions carefully and ensure you have the necessary approvals before proceeding.

## 1. Registering the Application in Azure

Access the Azure portal: [https://portal.azure.com](https://portal.azure.com)

To register your application, follow these steps:

1. Navigate to **App registrations** → **New registration**.
2. Provide a descriptive name for your application.
3. Specify the supported account type (e.g., **Accounts in this organizational directory only**).
4. Configure the redirect URI, if applicable.
5. Click **Register** to complete the registration process.

---

## 2. Configuring API Permissions for SharePoint

### Microsoft Graph Permissions

| Permission Name | Type      | Description                                | Admin Consent Required | Status    |
| --------------- | --------- | ------------------------------------------ | -------------------- | --------- |
| User.Read       | Delegated | Sign in and read the user profile          | No                   | Granted   |

### SharePoint Permissions

| Permission Name       | Type        | Description                                                            | Admin Consent Required | Status    |
| -------------------- | ----------- | ---------------------------------------------------------------------- | -------------------- | --------- |
| Sites.FullControl.All | Application | Full control over all site collections                                  | Yes                  | Granted   |
| Sites.Manage.All      | Application | Read and write items and lists across all site collections             | Yes                  | Granted   |
| Sites.Read.All        | Application | Read items across all site collections                                  | Yes                  | Granted   |
| Sites.ReadWrite.All   | Application | Read and write items across all site collections                        | Yes                  | Granted   |

**Important Notes:**

1. All **Application** type permissions require consent from an Azure AD administrator.  
2. The `User.Read` permission is **delegated**, meaning it executes on behalf of the signed-in user.  
3. Ensure that these permissions are granted **prior to executing the application** to avoid authorization errors when accessing SharePoint.

---

## 3. Exposing the API and Defining Scopes

Follow these steps to expose your API and configure scopes:

### Step 1: Access the Registered Application

1. Open the **Azure** portal.  
2. Navigate to **App registrations**.  
3. Select the target application (e.g., `mcp-sharepoint`).  

### Step 2: Expose an API

1. In the left-hand menu, select **Expose an API**.  
2. Confirm the API URI is configured (e.g., `api://xxxxxx-xxx-xxxxx-xxxxxx`).

### Step 3: Add a Scope

1. Click **+ Add a scope**.  
2. If this is your first scope, confirm the API URI prefix.  
3. Complete the form to define the new scope. The specific details are typically not critical.  
4. Click **Add scope** to save.

### Step 4: Authorize Client Applications (Optional)

1. To allow another application to consume your API, register it under **Authorized client applications**:  
   - Click **+ Add a client application**.  
   - Select the registered client application.  
   - Assign the permitted scopes.

---

## 4. Configuring Permissions in SharePoint

To enable the application to access SharePoint, register the required permissions using the SharePoint AppInv form. This method is valid until **April 2026**.

1. Access the SharePoint permissions form:  
 https://exampleorganization.sharepoint.com/sites/examplesite/_layouts/15/appinv.aspx

2. Enter the **Application ID** in the corresponding field and click **Lookup**. The application details will load.

3. In the **App Permission Request XML** field, enter the following XML:

```xml
<AppPermissionRequests AllowAppOnlyPolicy="true">
    <AppPermissionRequest Scope="http://sharepoint/content/sitecollection" Right="FullControl"/>
</AppPermissionRequests>
```
>    This XML grants full control over the site collection to the application. Adjust the settings as required based on your organizational needs.

    

4. Click Create or OK to save the permissions.

5. Upon completion, the application will have access to SharePoint according to the defined permissions, valid until April 2026.
