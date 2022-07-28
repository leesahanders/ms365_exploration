library(AzureAuth)
library(AzureGraph)
library(Microsoft365R)
library(shiny)
library(DT)
library(spsComps)

tenant = "b721d4a7-5012-49a6-8574-c2d52e59001b"
site_url = "https://rstudioinc.sharepoint.com/sites/integrations-testing"
app="29f4ba1e-3dd5-4f97-a94f-98c68497a9cc"
drive_name = "Documents"
file_src = "penguins_raw.csv"

# # - you should NEVER put secrets in code and instead do client_secret <- Sys.getenv("EXAMPLE_SHINY_CLIENT_SECRET", "")
client_secret <- Sys.getenv("EXAMPLE_SHINY_CLIENT_SECRET", "")

foo <- function() {
  message("Starting Microsoft Graph login and data pull")
  
  #The Azure Active Directory tenant for which to obtain a login client. Can be a name ("myaadtenant"), a fully qualified domain name ("myaadtenant.onmicrosoft.com" or "mycompanyname.com"), or a GUID. The default is to login via the "common" tenant, which will infer your actual tenant from your credentials.
  # Refer to: https://rdrr.io/github/Azure/AzureAuth/man/AzureR_dir.html
  create_AzureR_dir()
  
  # create a Microsoft Graph login
  # Refer to: https://cran.r-project.org/web/packages/Microsoft365R/vignettes/scripted.html
  gr <- create_graph_login(tenant, app, password=client_secret, auth_type="client_credentials")
  
  # Sharepoint site
  site <- gr$get_sharepoint_site(site_url)
  
  # Note the function get_drive() uses drive_id as a parameter (NOT drive name)
  drv <- site$get_drive(drive_name)
  
  # Downloading the file:
  drv$download_file(src = file_src, dest = "tmp.csv", overwrite = TRUE)
  
  # Reading file into memory (it would be a good idea to have handling here for detecting file type)
  data <- read.csv(file="tmp.csv", stringsAsFactors = FALSE, check.names=F)
  
  return(data)
}

#runApp(shinyApp(
ui = fluidPage(
  actionButton("btn1", "Log into Azure and Show Sharepoint File"),
  DT::dataTableOutput("table1")
)

server = function(input,output, session) {
  #Reactive values for for the data, https://mastering-shiny.org/index.html is a great resource for learning Shiny
  rv <- reactiveValues()
  rv$data <- NULL
  
  #btn
  observeEvent(input$btn1, {
    showNotification("btn1: Running")
    spsComps::shinyCatch({
      rv$data <- foo()
    },
    # blocking recommended
    blocking_level = "error",
    prefix = "My-project" #change console prefix if you don't want "SPS"
    )
  })
  
  # Output
  output$table1 <- DT::renderDataTable({
    rv$data
  })
}

# Run the application 
shinyApp(ui = ui, server = server)

#### The AzureR packages save your login sessions so that you don’t need to reauthenticate each time. If you’re experiencing authentication failures, you can try clearing the saved data by running the following code:
#   
# AzureAuth::clean_token_directory()
# AzureGraph::delete_graph_login(tenant="mytenant")