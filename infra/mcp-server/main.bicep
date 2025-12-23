targetScope = 'subscription'

@minLength(1)
@maxLength(64)
@description('Name of the environment that can be used as part of naming resource convention')
param environmentName string

@minLength(1)
@description('Primary location for all resources')
param location string

@description('Azure AD Tenant ID for authentication')
param azureTenantId string = ''

@description('Azure AD Client ID (App Registration) for Microsoft Graph API')
param azureClientId string = ''

@description('Azure AD Client Secret for Microsoft Graph API')
@secure()
param azureClientSecret string = ''

// Optional parameters
param containerAppName string = ''
param containerRegistryName string = ''
param logAnalyticsName string = ''
param applicationInsightsName string = ''

var abbrs = loadJsonContent('../abbreviations.json')
var resourceToken = toLower(uniqueString(subscription().id, environmentName, location))
var tags = { 'azd-env-name': environmentName }

// Resource group
resource rg 'Microsoft.Resources/resourceGroups@2022-09-01' = {
  name: 'rg-${environmentName}'
  location: location
  tags: tags
}

// Log Analytics workspace
module logAnalytics 'br/public:avm/res/operational-insights/workspace:0.3.4' = {
  name: 'logAnalytics'
  scope: rg
  params: {
    name: !empty(logAnalyticsName) ? logAnalyticsName : '${abbrs.operationalInsightsWorkspaces}${resourceToken}'
    location: location
    tags: tags
  }
}

// Application Insights
module applicationInsights 'br/public:avm/res/insights/component:0.3.0' = {
  name: 'applicationInsights'
  scope: rg
  params: {
    name: !empty(applicationInsightsName) ? applicationInsightsName : '${abbrs.insightsComponents}${resourceToken}'
    location: location
    tags: tags
    workspaceResourceId: logAnalytics.outputs.resourceId
  }
}

// Container Registry
module containerRegistry 'br/public:avm/res/container-registry/registry:0.1.1' = {
  name: 'containerRegistry'
  scope: rg
  params: {
    name: !empty(containerRegistryName) ? containerRegistryName : '${abbrs.containerRegistryRegistries}${resourceToken}'
    location: location
    tags: tags
    acrSku: 'Basic'
    acrAdminUserEnabled: true
  }
}

// Container Apps Environment
module containerAppsEnvironment 'br/public:avm/res/app/managed-environment:0.4.5' = {
  name: 'containerAppsEnvironment'
  scope: rg
  params: {
    name: '${abbrs.appManagedEnvironments}${resourceToken}'
    location: location
    tags: tags
    logAnalyticsWorkspaceResourceId: logAnalytics.outputs.resourceId
    zoneRedundant: false
  }
}

// Container App for MCP Server
module containerApp 'container-app.bicep' = {
  name: 'containerApp'
  scope: rg
  params: {
    name: !empty(containerAppName) ? containerAppName : '${abbrs.appContainerApps}mcp-${resourceToken}'
    location: location
    tags: union(tags, { 'azd-service-name': 'mcp-server' })
    containerAppsEnvironmentId: containerAppsEnvironment.outputs.resourceId
    containerRegistryName: containerRegistry.outputs.name
    applicationInsightsConnectionString: applicationInsights.outputs.connectionString
    azureTenantId: !empty(azureTenantId) ? azureTenantId : tenant().tenantId
    azureClientId: azureClientId
    azureClientSecret: azureClientSecret
  }
}

// Outputs
output AZURE_LOCATION string = location
output AZURE_TENANT_ID string = tenant().tenantId
output AZURE_RESOURCE_GROUP_NAME string = rg.name
output AZURE_CONTAINER_REGISTRY_ENDPOINT string = containerRegistry.outputs.loginServer
output AZURE_CONTAINER_REGISTRY_NAME string = containerRegistry.outputs.name
output AZURE_CONTAINER_APPS_ENVIRONMENT_NAME string = containerAppsEnvironment.outputs.name
output AZURE_CONTAINER_APP_NAME string = containerApp.outputs.name
output AZURE_CONTAINER_APP_FQDN string = containerApp.outputs.fqdn
output MCP_ENDPOINT string = 'https://${containerApp.outputs.fqdn}/mcp'
output APPLICATIONINSIGHTS_CONNECTION_STRING string = applicationInsights.outputs.connectionString
