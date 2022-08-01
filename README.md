# ms365_exploration

> :warning: This is a work in progress with various notes on trying to connect the RStudio Pro products with data stored in Sharepoint. Please come back later for this to be updated once I (hopefully) understand something worth sharing. It is possible (likely) that I am misunderstanding capabilities listed below so read at your own risk. Also anything actually well written was probably taken from the documentation from one of the amazing packages I'm playing with - credit where credit is due (Microsoft365R, rsconnect, pins, etc).


## Background 

Use of Microsoft products at the enterprise level is not uncommon. However getting that stored data from the Microsoft online resources into the RStudio IDE and pro products is different from the general recommended approaches of storing that information in databases. Below this article goes through the steps with copious examples using Sharepoint Online for how to:

1.  Authenticate to Microsoft (Azure)
2.  Pull and push data

The distinction to emphasize is the approach options for interactive content (where a user / viewer is available for pass-through authentication or interactive authentication) versus the approach options for non-interactive content (where content is being run without a user/ viewer available or a user is not available for authenticating to the service). 

## Outline

The below outline may be useful for scoping which options are appropriate for particular use cases, and deciding on a method to use within your organization or particular piece of content being developed. 

-   Microsoft365R: A user sign-in flow (via redirect url and user permissions, best for interactive applications such as Workbench)
-   Microsoft365R: Embedded user / service account credentials (via embedding account username and password and user permissions, can be used in interactive and non-interactive contexts)
-   Microsoft365R: Embedded authentication (via client service and application permissions, can be used in interactive and non-interactive contexts)
-   Mapping as a network drive: last resort option
-   Pins: Useful for creating new pinned caches of data using existing Microsoft online resources for storage and easy retrieval later

Before diving into discussion and examples for the different authentication options there are a few cases not covered in this article that may be useful to list here, as well as recommended resources, for interested readers. 

-   For user level authentication into servers / applications refer to the [Marketplace offering](https://azuremarketplace.microsoft.com/en-us/marketplace/apps/aad.rstudioconnect?tab=Overview) and the [Connect documentation](<https://docs.rstudio.com/connect/admin/authentication/saml-based/okta-saml/#idp-config>). 
-   For python users the [Microsoft REST API](https://github.com/vgrem/Office365-REST-Python-Client) is the official Microsoft supported method with examples located [here](https://github.com/vgrem/Office365-REST-Python-Client/tree/master/examples/sharepoint/files).

### **Authentication: Microsoft365R**

The Microsoft supported method for authentication is through use of the [Microsoft365R](https://github.com/Azure/Microsoft365R) package which was developed by Hong Ooi. 

Documentation for Microsoft365R is very thorough and can be accessed at [1](<https://github.com/Azure/Microsoft365R>) and [2](<https://cran.r-project.org/web/packages/Microsoft365R/>) with the scopes detailed [here]( <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md>). Microsoft365R was announced in the [Community forums](https://community.rstudio.com/t/microsoft365r-interface-to-microsoft-365-sharepoint-onedrive-etc/94287) and additional useful discussion can be found there as well as other users of the package. 

Authentication to Microsoft is handled through Microsoft's Azure cloud platform ('Azure Active Directory') with the creation of an application and assigning different levels of permissions in order to obtain 'Oath' 2.0 tokens. Broadly speaking the authentication options can be split into three approaches:

-   A user sign-in flow (via redirect url and user permissions, or with a device code needing user permissions and the enabling of mobile and desktop flows)
-   Service principal/Client secret Embedded authentication (via client service and application permissions)
-   Embedded user / service account credentials (via embedding account username and password and user permissions)

It is worth noting that for the options listed above support from the Microsoft administrator will likely be needed (for creating applications, adding user/application permissions, redirect URI's).

#### **User Sign-In Authentication**

User level permissions will need to be assigned as detailed in <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md> with additional Redirect URI's added as needed. It’s possible to set more than one Redirect URI per app, so a single app registration can be used for multiple platforms or pieces of content. Adding / modifying the Redirect URI is done through Azure under `App Registrations` -\> `select your app` -\> `Authentication` -\> `Platform configurations` -\> `Mobile and desktop applications` -\>`Add URI` and also requires enabling nativeclient.

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

Alternatively the device code flow can be used. This is especially useful in cases where adding a Redirect URI isn't possible or where it is easier for the user to open a new tab and enter a code provided in the console for authentication. The process for enabling this workflow is though the App Registration dashboard in Azure -\> `click on the created app` -\> `Authentication` -\> `Allow public client flows` and set `Enable the following mobile and desktop flows` to `yes`.

Example:

    library(Microsoft365R)

    site_url = MySharepointSiteURL
    app = MyApp

    site <- get_sharepoint_site(site_url = site_url, app=app, auth_type="device_code")

#### **Service principal/Client secret Embedded authentication**

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

User level permissions will need to be assigned as detailed in <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md>.

The Service principal/Client secret Embedded authentication approach discussed above is the Microsoft recommended approach for unattended scripts. However that method may not support all cases and also requires additional access that could make gaining the needed permissions from local Microsoft administrators more challenging.

This method does not require application level permissions and instead uses a user account or a service account for gaining access via a scripted command. Username / password should be embedded as environmental variables so that they are never exposed in the code directly. Starting with version 1.6, RStudio Connect allows [Environment Variables](https://docs.rstudio.com/connect/admin/security-and-auditing/#application-environment-variables) to be saved at the application level. The variables are encrypted on-disk, and in-memory.

-   This can be done at the application level with [deployment](https://db.rstudio.com/best-practices/deployment/) through the [Connect UI](https://support.rstudio.com/hc/en-us/articles/228272368-Managing-your-content-in-RStudio-Connect) or at the [server level with support from the Connect administrator](https://support.rstudio.com/hc/en-us/articles/360016606613-Environment-variables-on-RStudio-Connect)
-   Additional Microsoft supported R packages are useful, as shown in the below example, in order to remove any interactive elements when calling functions.

An additional recommendation is to create a service account rather than using a users username and password for authentication. This could be done as a user or by your administrator. The discussion on the [Using Microsoft365R in an unattended script](https://cran.r-project.org/web/packages/Microsoft365R/vignettes/scripted.html) vignette articulates clearly the key points and I recommend reading the article. In addition to the points made in the vignette it is worth adding that having a service account dedicated to a specific content application enables rapid troubleshooting as an administrator or a publisher of multiple pieces of content that can otherwise be much more challenging. 

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

#### Troubleshooting authentication failures

In the case of authentication failures clearing cached authentication tokens/files can be done with:

    AzureAuth::clean_token_directory()
    AzureGraph::delete_graph_login(tenant="mytenant")


### **Authentication: Mapping Sharedrive as a Network Drive**

As a last resort mapping Sharedrive as a network drive to the hosting server could be considered, using a program such as expandrive. The resources below may be of interest, however please note that this method was not tested nor is endorsed by RStudio. 

 - <https://www.clouddirect.net/knowledge-base/KB0011543/mapping-a-sharepoint-site-as-a-network-drive>
 - Mapping Sharedrive as a Network Drive: <https://www.clouddirect.net/knowledge-base/KB0011543/mapping-a-sharepoint-site-as-a-network-drive> and <https://www.expandrive.com/onedrive-for-linux/> specifically this version <https://docs.expandrive.com/server-edition/installation> with the documentation at <https://docs.expandrive.com/server-edition/installation>


### Pull and push data

There are great options for navigating through the sharepoint site and making the data available to other developers or to content that is hosted. Content and publishers can explore and make data available for consumption after authentication either through the Microsoft365R functions or content can be created with the [board_ms365() function from the pins](https://pins.rstudio.com/reference/board_ms365.html) package and then accessed.

Example with [Microsoft365R](https://cran.r-project.org/web/packages/Microsoft365R/vignettes/od_sp.html):

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



Microsoft resources can be used for hosting content in pins format that can then be accessed programmatically and manually. Note that this does require the data to be stored in a pins compatible format, meaning that prior existing data may not be able to be accessed using this method. 

Example with [board_ms365() from pins](https://pins.rstudio.com/reference/board_ms365.html):

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

