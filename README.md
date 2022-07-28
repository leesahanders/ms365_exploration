# ms365_exploration

> :warning: This is a work in progress with various notes on trying to connect the RStudio Pro products with data stored in Sharepoint. All notes below are a stream of consciousness, please come back later for this to be updated once I (hopefully) understand something worth sharing. It is possible (likely) that I am misunderstanding capabilities listed below so read at your own risk. Also anything actually well written was probably taken from the documentation from one of the amazing packages I'm playing with - credit where credit is due (Microsoft365R, rsconnect, pins, etc). 

The critical distinction to emphasize is the approach options for interactive content (where a user / viewer is available for pass-through authentication or interactive authentication) versus the approach options for non-interactive content (where content is being run without a user/ viewer available or a user is not available for authenticating to the service).

For deep dives into the various approaches the below resources have been super useful:

-   Pins: <https://pins.rstudio.com/reference/board_ms365.html>
-   Mapping Sharedrive as a Network Drive: <https://www.clouddirect.net/knowledge-base/KB0011543/mapping-a-sharepoint-site-as-a-network-drive>
-   Microsoft365R: <https://github.com/Azure/Microsoft365R> and <https://cran.r-project.org/web/packages/Microsoft365R/>

For user authentication into systems refer to the Marketplace offering at <https://azuremarketplace.microsoft.com/en-us/marketplace/apps/aad.rstudioconnect?tab=Overview> and the Connect documentation at <https://docs.rstudio.com/connect/admin/authentication/saml-based/okta-saml/#idp-config>.

### **Mapping Sharedrive as a Network Drive**

Upcoming!

<https://www.clouddirect.net/knowledge-base/KB0011543/mapping-a-sharepoint-site-as-a-network-drive>

### **Content Level Access Through Microsoft365R**

Additional documentation for Microsoft365R is very thorough and can be accessed at <https://github.com/Azure/Microsoft365R> and <https://cran.r-project.org/web/packages/Microsoft365R/> with the scopes detailed here: <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md>

Broadly speaking the authentication methodology can be split into two approaches:

-   A user sign-in flow (via redirect url and user permissions)
-   Embedded authentication (via client service and application permissions)

Sharepoint Online is part of the Microsoft 365 ecosystem with access controlled through Azure as an Application. While the scope of this article is specific to data stored in Sharepoint Online this approach and package supports the entire Microsoft Ecosystem including Outlook, Teams, etc.

It is worth noting that for both options support from the Microsoft administrator will likely be needed (for creating applications, adding user/application permissions, redirect URI's).

#### **User Sign-In Authentication**

User level permissions will need to be assigned as detailed in <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md> with additional Redirect URI's added as needed. It’s possible to set more than one Redirect URI per app, so a single app registration can be used for multiple platforms or pieces of content. Adding / modifying the Redirect URI is done through Azure under `App Registrations` -\> `select your app` -\> `Authentication` -\> `Platform configurations` -\> `Mobile and desktop applications` -\>`Add URI` and also requires enabling nativeclient.

-   For the desktop RStudio IDE the URI is: `http://localhost:1410/`.
-   Azure requires SSL certificates for Redirect URI's for non-local connections. This means that for connecting via a server (for example for Connect and Workbench) the connection should be HTTPS in order to be added then should work as expected.
    -   In some scenarios, for example for Connect hosted applications or for Workbench, instead of adding the Redirect URI for each distinct URL a wildcard can be added to enable access for the entire server. There are restrictions on this approach as detailed in the Microsoft documentation as listed at: <https://docs.microsoft.com/en-us/azure/active-directory/develop/reply-url#restrictions-on-wildcards-in-redirect-uris>
-   For an app hosted in shinyapps.io this would be a URL of the form <https://youraccount.shinyapps.io/appname> (including the special port number if specified).

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

#### **Embedded authentication**

Content in a non-interactive context (IE scheduled content for example) won't have a user account available for authentication. In this case using an embedded authentication workflow will be necessary. There are several approaches outlined in <https://cran.r-project.org/web/packages/Microsoft365R/vignettes/scripted.html> but as using a Service Principal via using a Client Secret is the Microsoft recommended approach that will be the approach detailed in this article.

The Azure application being used for access needs application level permissions. The permissions can be based off of the user permissions documented at <https://github.com/Azure/Microsoft365R/blob/master/inst/app_registration.md> but can be assigned as needed for the application and to comply with any restrictions from the IT administration.

Application permissions are more powerful than user permissions so it is important to emphasize that exposing the client secret directly should be avoided. Instead adding the client secret through an environmental variable is recommended. Starting with version 1.6, RStudio Connect allows [Environment Variables](https://docs.rstudio.com/connect/admin/security-and-auditing/#application-environment-variables) to be saved at the application level. The variables are encrypted on-disk, and in-memory.

-   This can be done at the application level with [deployment](https://db.rstudio.com/best-practices/deployment/) through the [Connect UI](https://support.rstudio.com/hc/en-us/articles/228272368-Managing-your-content-in-RStudio-Connect) or at the [server level with support from the Connect administrator](https://support.rstudio.com/hc/en-us/articles/360016606613-Environment-variables-on-RStudio-Connect)
-   Additional Microsoft supported R packages might be useful, as shown in the below example, in order to remove any interactive elements when calling functions.

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

In the case of authentication failures then clearing cached authentication tokens/files can be done with:

    AzureAuth::clean_token_directory()
    AzureGraph::delete_graph_login(tenant="mytenant")

### After authentication how do developers find their data?

There are great options for navigating through the sharepoint site and making the data available to other developers or to content that is hosted. Content and publishers can explore and make data available for consumption after authentication either through the Microsoft365R functions or content can be created with the [board_ms365() function from the pins](https://pins.rstudio.com/reference/board_ms365.html) package and then accessed.

Example with [Microsoft365R](https://cran.r-project.org/web/packages/Microsoft365R/vignettes/od_sp.html) :

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

Example with [board_ms365() from pins](https://pins.rstudio.com/reference/board_ms365.html):

Microsoft resources can be used for hosting content in pins format that can then be accessed programmatically and manually.

(I don't think this allows for access to existing content within Microsoft unless it has been saved using the pin format -so it wouldn't be a good choice for say if users are uploading files to Sharepoint manually and then there is a programmatic script parsing them to reupload an aggregate/analysis. )

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
