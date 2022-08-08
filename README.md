# ms365_exploration

> :warning: This is a work in progress for understanding the options for connecting the RStudio Pro products with data stored in Sharepoint Online (not server). This is a personal repo and as such might change without notice (though if you needed anything feel free to email me at [lisa.anders\@rstudio.com](mailto:lisa.anders@rstudio.com){.email} and I'll do my best to help recover it). Huge thanks to Hong Ooi and his fantastic documentation on Microsoft365R.

## Background

Use of Microsoft products at the enterprise level is common. Getting that stored data from the Microsoft online resources into the RStudio IDE and pro products uses packages wrapping the Microsoft system API for interfacing.

## Introduction

Microsoft 365 is a subscription extension of the Microsoft Office product line with cloud hosting support. Microsoft 365 uses Azure Active Directory (Azure AD) for user authentication and application access through developed API's. The Microsoft supported method for interfacing with R developed content is with the [Microsoft365R](https://github.com/Azure/Microsoft365R) package which was developed by Hong Ooi and has extensive documentation.

## Summary

There are three main authentication approaches supported by [Microsoft365R](https://github.com/Azure/Microsoft365R). Note that multiple approaches can be supported at the same time.

| **Method**                            | **auth_type**        | **Privileges** | **Capability**                                                         |
|---------------------|-----------------|-----------------|-----------------|
| **User sign-in flow**                 | device_code, default | User           | Interactive only (local IDE and Workbench, interactive Shiny content)  |
| **Service principal / Client secret** | client_credentials   | Application    | Interactive and non-interactive (same as above plus scheduled content) |
| **Embedded credentials**              | resource_owner       | User           | Interactive and non-interactive (same as above plus scheduled content) |

Authentication for [Microsoft365R](https://github.com/Azure/Microsoft365R) is through Microsoft's Azure cloud platform through a registered [application](https://docs.microsoft.com/en-us/azure/active-directory/develop/app-objects-and-service-principals) with [appropriate assigned permissions](https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md) in order to obtain ['OAuth 2.0' tokens](https://github.com/Azure/AzureAuth).

Depending on your organizations security policy some steps may require support from your Azure Global Administrator.

### Administration Overview

**User sign-in flow**

A custom app can be created or the default app registration `d44a05d5-c6a5-4bbb-82d2-443123722380` that comes with the [Microsoft365R](https://github.com/Azure/Microsoft365R) package can be used. The user permissions will need to be enabled as specified in [the app registrations page](https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md). Depending on your organizations security policy access to your tenant may need to be granted by an Azure Global Administrator. Additionally Redirect URLs will need to be added through Azure under `App Registrations` -\> `select your app` -\> `Authentication` -\> `Platform configurations` -\> `Mobile and desktop applications` -\>`Add URI` as well as also enabling `nativeclient`.

For adding Redirect URLs, which will give a typical web-app authentication experience for interactive applications:

-   For the desktop RStudio IDE the URL is: `http://localhost:1410/`.
-   For content hosted in shinyapps.io this would be of the form `https://youraccount.shinyapps.io/appname` (including the port number if specified).
-   Typically a SSL certificate will be required for non-local connections, including for Microsoft Azure. This means that the Connect and Workbench URLs will need to be HTTPS. A wildcard could be used instead of adding the Redirect URL for each piece of content/user where appropriate for server-wide access.

Enabling the device code workflow is though the App Registration dashboard in Azure -\> `click on the created app` -\> `Authentication` -\> `Allow public client flows` and set `Enable the following mobile and desktop flows` to `yes`.

**Service principal / Client secret**

A custom app will need to be registered in Azure with Application permissions. The permissions can be based off of the [user permissions](https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md) but can be assigned as needed for the application and to comply with any security restrictions.

Application permissions are more powerful than user permissions so it is important to emphasize that exposing the client secret directly should be avoided. As a control using environmental variable's for storing the client secret is recommended. Starting with version 1.6, RStudio Connect allows [Environment Variables](https://docs.rstudio.com/connect/admin/security-and-auditing/#application-environment-variables) to be saved at the application level. The variables are encrypted on-disk, and in-memory.

-   This can be done at the application level with [deployment](https://db.rstudio.com/best-practices/deployment/) through the [Connect UI](https://support.rstudio.com/hc/en-us/articles/228272368-Managing-your-content-in-RStudio-Connect).

**Embedded credentials**

A custom app can be created or the default app registration `d44a05d5-c6a5-4bbb-82d2-443123722380` that comes with the [Microsoft365R](https://github.com/Azure/Microsoft365R) package can be used. The user permissions will need to be enabled as specified in [the app registrations page](https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md). Depending on your organizations security policy access to your tenant may need to be granted by an Azure Global Administrator.

 - Working with your Microsoft administrator to create service accounts per content can be useful to enable fast troubleshooting and easier collaboration on content with multiple developers is recommended. 

## Authentication Examples

The Microsoft supported method for authentication is through use of the [Microsoft365R](https://github.com/Azure/Microsoft365R) package which was developed by Hong Ooi.

Documentation for Microsoft365R is very thorough and can be accessed at [1](https://github.com/Azure/Microsoft365R) and [2](https://cran.r-project.org/web/packages/Microsoft365R/) with the scopes detailed [here](https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md). Microsoft365R was announced in the [Community forums](https://community.rstudio.com/t/microsoft365r-interface-to-microsoft-365-sharepoint-onedrive-etc/94287) and additional useful discussion can be found there as well as other users of the package.

Authentication to Microsoft is handled through Microsoft's Azure cloud platform ('Azure Active Directory') with the creation of an application and assigning different levels of permissions in order to obtain 'Oath' 2.0 tokens. Broadly speaking the authentication options can be split into three approaches:

-   A user sign-in flow (via redirect url and user permissions, or with a device code needing user permissions and the enabling of mobile and desktop flows)
-   Service principal/Client secret Embedded authentication (via client service and application permissions)
-   Embedded user / service account credentials (via embedding account username and password and user permissions)

It is worth noting that for the options listed above support from the Microsoft administrator will likely be needed (for creating applications, adding user/application permissions, redirect URI's).

### **User Sign-In Authentication: Default method**

User level permissions will need to be assigned as detailed in <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md> with additional Redirect URI's added as needed. It's possible to set more than one Redirect URI per app, so a single app registration can be used for multiple platforms or pieces of content. Adding / modifying the Redirect URI is done through Azure under `App Registrations` -\> `select your app` -\> `Authentication` -\> `Platform configurations` -\> `Mobile and desktop applications` -\>`Add URI` and also requires enabling nativeclient.

-   For the desktop RStudio IDE the URI is: `http://localhost:1410/`.
-   Azure requires SSL certificates for Redirect URI's for non-local connections. This means that for connecting via a server (for example for Connect and Workbench) the connection should be HTTPS in order to be added then should work as expected.
    -   In some scenarios, for example for Connect hosted applications or for Workbench, instead of adding the Redirect URI for each distinct URL a wildcard can be added to enable access for the entire server. There are restrictions on this approach as detailed in the Microsoft documentation as listed at: <https://docs.microsoft.com/en-us/azure/active-directory/develop/reply-url#restrictions-on-wildcards-in-redirect-uris>
-   For an app hosted in shinyapps.io this would be a URL of the form <https://youraccount.shinyapps.io/appname> (including the port number if specified).

Example:

    library(Microsoft365R)

    site_url = MySharepointSiteURL
    app = MyApp
    drive_name = MyDrive # For example by default this will likely be "Documents"
    file_src = MyFileName.TheExtension
    file_dest = MyFileNameDestination.TheExtension

    site <- get_sharepoint_site(site_url = site_url, app = app)

    # View the sharepoint sites that we have access to
    list_sharepoint_sites(app=app)

    # Note the function get_drive() uses drive_id as a parameter (NOT drive name)
    drv <- site$get_drive(drive_name)

    # Here we can retrieve lists of the different types of items in our sharepoint site. Documents uploaded under 'Documents' are retrieved with list_files() 
    drv$list_items()
    drv$list_files() 
    drv$list_shared_files()
    drv$list_shared_items()

    # Open a file in a web browser tab: 
    drv$open_item(file_src)

    # Downloading and uploading a file (in this example a csv)
    drv$download_file(src = file_src, dest = file_dest, overwrite = TRUE)
    data = read.csv(file_dest)
    drv$upload_file(src = file_dest, dest = file_dest)

#### User sign-in flow: Device code workflow

Alternatively the device code flow can be used. This is especially useful in cases where adding a Redirect URI isn't possible or where it is easier for the user to open a new tab and enter a code provided in the console for authentication. The process for enabling this workflow is though the App Registration dashboard in Azure -\> `click on the created app` -\> `Authentication` -\> `Allow public client flows` and set `Enable the following mobile and desktop flows` to `yes`.

Example:

    library(Microsoft365R)

    site_url = MySharepointSiteURL
    app = MyApp

    site <- get_sharepoint_site(site_url = site_url, app=app, auth_type="device_code")

### **Service principal/Client secret Embedded authentication**

Content in a non-interactive context (IE scheduled content for example) won't have a user account available for authentication. There are several approaches outlined in <https://cran.r-project.org/web/packages/Microsoft365R/vignettes/scripted.html>, with the Service Principal via using a Client Secret discussed in this section being the Microsoft recommended approach.

The Azure application being used for access needs application level permissions. The permissions can be based off of the user permissions documented at <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md> but can be assigned as needed for the application and to comply with any restrictions from the IT administration.

Application permissions are more powerful than user permissions so it is important to emphasize that exposing the client secret directly should be avoided. Instead adding the client secret through an environmental variable is recommended. Starting with version 1.6, RStudio Connect allows [Environment Variables](https://docs.rstudio.com/connect/admin/security-and-auditing/#application-environment-variables) to be saved at the application level. The variables are encrypted on-disk, and in-memory.

-   This can be done at the application level with [deployment](https://db.rstudio.com/best-practices/deployment/) through the [Connect UI](https://support.rstudio.com/hc/en-us/articles/228272368-Managing-your-content-in-RStudio-Connect) or at the [server level with support from the Connect administrator](https://support.rstudio.com/hc/en-us/articles/360016606613-Environment-variables-on-RStudio-Connect)
-   Additional Microsoft supported R packages are useful, as shown in the below example, in order to remove any interactive elements when calling functions.

Example:

    # This example is used in non-interactive content such as a RMarkdown or in interactive content where the app is handling authentication for example with certain Shiny apps

    library(AzureAuth)
    library(AzureGraph)
    library(Microsoft365R)

    tenant = MyTenant
    site_url = MySharepointSiteURL
    app = MyApp
    drive_name = MyDrive # For example by default this will likely be "Documents"
    file_src = MyFileName.TheExtension
    file_dest = MyFileNameDestination.TheExtension

    # You should NEVER put secrets in code and instead do
    client_secret <- Sys.getenv("EXAMPLE_SHINY_CLIENT_SECRET", "")

    DownloadSharepointFile <- function() {
      message("Microsoft Graph login and data pull")
      
      create_AzureR_dir()
      message(AzureR_dir())
      
      # create a Microsoft Graph login
      gr <- create_graph_login(tenant, app, password=client_secret, auth_type="client_credentials")

      # Sharepoint site
      site <- gr$get_sharepoint_site(site_url)

      # Note the function get_drive() uses drive_id as a parameter (NOT drive name)
      drv <- site$get_drive(drive_name)

      # Downloading the file:
      drv$download_file(src = file_src, dest = file_dest, overwrite = TRUE)

      # The data could then be read back in for visualization
    }

    data <- DownloadSharepointFile()

#### **Embedded user / service account credentials**

The Service principal/Client secret Embedded authentication approach discussed above is the Microsoft recommended approach for unattended scripts. However that method may not support all cases and also requires additional access that could make gaining the needed permissions from local Microsoft administrators more challenging.

Instead signing in by embedding a service account or user's credentials may be used. This method does not require application level permissions for gaining access via a scripted command. Sensitive variables such as Username / password should be embedded as environmental variables so that they are never exposed in the code directly. Starting with version 1.6, RStudio Connect allows [Environment Variables](https://docs.rstudio.com/connect/admin/security-and-auditing/#application-environment-variables) to be saved at the application level. The variables are encrypted on-disk, and in-memory.

-   This can be done at the application level with [deployment](https://db.rstudio.com/best-practices/deployment/) through the [Connect UI](https://support.rstudio.com/hc/en-us/articles/228272368-Managing-your-content-in-RStudio-Connect) or at the [server level with support from the Connect administrator](https://support.rstudio.com/hc/en-us/articles/360016606613-Environment-variables-on-RStudio-Connect)

-   Working with your Microsoft administrator to create service accounts per content can be useful to enable fast troubleshooting and easier collaboration on content with multiple developers. The discussion on the [Using Microsoft365R in an unattended script](https://cran.r-project.org/web/packages/Microsoft365R/vignettes/scripted.html) vignette articulates clearly the considerations when using this approach.

-   Use of the Microsoft developed package [AzureAuth](https://github.com/Azure/AzureAuth) may be needed for fully removing console prompt elements so a script can be run in a non-interactive context, for example by explicitly defining the token directory with `AzureAuth::create_AzureR_dir()`.

-   User level permissions will need to be assigned as detailed in in the 'User sign-in flow' sections above.

Example:

        # This example is used in non-interactive content such as a RMarkdown or in interactive content where the app is handling authentication for example with certain Shiny apps

        library(AzureAuth)
        library(AzureGraph)
        library(Microsoft365R)

        tenant = MyTenant
        site_url = MySharepointSiteURL
        app = MyApp
        drive_name = MyDrive # For example by default this will likely be "Documents"
        file_src = MyFileName.TheExtension
        file_dest = MyFileNameDestination.TheExtension

        # You should NEVER put passwords in code and instead do
        user <- Sys.getenv("EXAMPLE_MS365R_SERVICE_USER")
        pwd <- Sys.getenv("EXAMPLE_MS365R_SERVICE_PASSWORD")

        DownloadSharepointFile <- function() {
          message("Microsoft Graph login and data pull")
          
          create_AzureR_dir()
          message(AzureR_dir())
          
          # create a Microsoft Graph login
          gr <- create_graph_login(tenant, app, 
                             username = user, 
                             password = pwd,
                             auth_type="resource_owner")

          # Sharepoint site
          site <- gr$get_sharepoint_site(site_url)

          # Note the function get_drive() uses drive_id as a parameter (NOT drive name)
          drv <- site$get_drive(drive_name)

          # Downloading the file:
          drv$download_file(src = file_src, dest = file_dest, overwrite = TRUE)

          # The data could then be read back in for visualization
        }

        data <- DownloadSharepointFile()

### Troubleshooting authentication failures

In the case of authentication failures clearing cached authentication tokens/files can be done with:

    AzureAuth::clean_token_directory()
    AzureGraph::delete_graph_login(tenant="mytenant")

## Sharepoint Examples

### Microsoft365R

The authentication method used in this example could be swapped out for any of the examples shown above. The documentation on [Microsoft365R](https://github.com/Azure/Microsoft365R) contains extensive examples beyond what is included below.

    library(Microsoft365R)
    library(AzureAuth)

    site_url = MySharepointSiteURL
    tenant = MyTenant
    app = MyApp
    drive_name = MyDrive # For example by default this will likely be "Documents"
    file_src = MyFileName.TheExtension
    file_dest = MyFileNameDestination.TheExtension

    # Add sensitive variables as environmental variables so they aren't exposed
    client_secret <- Sys.getenv("EXAMPLE_SHINY_CLIENT_SECRET", "")

    # Create auth token cache directory, otherwise it will prompt the user on the console for input
    create_AzureR_dir()

    # create a Microsoft Graph login
    gr <- create_graph_login(tenant, app, password=client_secret, auth_type="client_credentials")

    # An example of using the Graph login to connect to a Sharepoint site
    site <- gr$get_sharepoint_site(site_url)

    # Note the function get_drive() uses drive_id as a parameter (NOT drive name)
    drv <- site$get_drive(drive_name)

    # Download a specific file
    drv$download_file(src = file_src, dest = "tmp.csv", overwrite = TRUE)

    # Retrieve lists of the different types of items in our sharepoint site. Documents uploaded under 'Documents' are retrieved with list_files(). 
    drv$list_items()
    drv$list_files() 
    drv$list_shared_files()
    drv$list_shared_items()

    # Files can also be uploaded
    drv$upload_file(src = file_dest, dest = file_dest)

### Pins

Microsoft resources can be used for hosting content in pins format using [board_ms365() from pins](https://pins.rstudio.com/reference/board_ms365.html).

    library(Microsoft365R)
    library(pins)

    site_url = MySite
    app=MyApp

    site <- get_sharepoint_site(site_url = site_url, app=app, auth_type="device_code")

    doclib <- site$get_drive()

    # Connect ms365 as a pinned board. If this folder doesn't already exist it will be created. 
    board <- board_ms365(drive = doclib, "general/project1/board")

    # Write a dataset as a pin to Sharepoint
    board %>% pin_write(iris, "iris", description = "This is a test")

    # View the metadat of the pin we just created 
    board %>% pin_meta("iris")

    # Read a pin
    test <- board %>% pin_read("iris")

## Other Microsoft Related Resources

There are a few cases not covered in this article where the below resources may be useful:

-   For user level authentication into servers refer to the [Marketplace offering](https://azuremarketplace.microsoft.com/en-us/marketplace/apps/aad.rstudioconnect?tab=Overview) and the [Connect documentation](https://docs.rstudio.com/connect/admin/authentication/saml-based/okta-saml/#idp-config).

-   For python users the [Microsoft REST API](https://github.com/vgrem/Office365-REST-Python-Client) is the Microsoft developed method with [examples](https://github.com/vgrem/Office365-REST-Python-Client/tree/master/examples/sharepoint/files).

-   As a last resort mapping Sharedrive, OneNote, or other systems as a network drive to the hosting server could be considered, using a program such as [expandrive](https://www.expandrive.com/onedrive-for-linux/).

## End

On the off chance that anyone makes it to the end this article got a chuckle out of me and may be relatable: <https://www.theregister.com/2022/07/15/on_call/>
