library(AzureAuth)
library(AzureGraph)
library(Microsoft365R)
library(shiny)
library(DT)
library(spsComps)

tenant = Sys.getenv("EXAMPLE_TENANT", "")
app=Sys.getenv("EXAMPLE_APP", "")
site_url = "https://rstudioinc.sharepoint.com/sites/integrations-testing"
drive_name = "Documents" 
file_src = "penguins_raw.csv"

# Add sensitive variables as environmental variables so they aren't exposed
client_secret <- Sys.getenv("EXAMPLE_SHINY_CLIENT_SECRET", "")

# Create auth token cache directory, otherwise it will prompt the user on the console for input
create_AzureR_dir()

foo <- function() {
  message("Microsoft Graph login and data pull")
  
  # Create a Microsoft Graph login
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


ui = fluidPage(
  actionButton("btn1", "Log into Azure and Show Sharepoint File"),
  DT::dataTableOutput("table1")
)

server = function(input,output, session) {
  #Reactive values for for the data, https://mastering-shiny.org/index.html is a great resource for learning Shiny
  rv <- reactiveValues()
  rv$data <- NULL
  
  #btn1
  observeEvent(input$btn1, {
    showNotification("btn1: Running")
    spsComps::shinyCatch({
      rv$data <- foo()
    },
    blocking_level = "error",
    prefix = "My-project" 
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