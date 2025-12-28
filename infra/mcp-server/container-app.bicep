@description('Name of the Container App')
param name string

@description('Location for the resource')
param location string = resourceGroup().location

@description('Tags for the resource')
param tags object = {}

@description('Container Apps Environment ID')
param containerAppsEnvironmentId string

@description('Container Registry name')
param containerRegistryName string

@description('Container Registry login server')
param containerRegistryLoginServer string

@description('Application Insights connection string')
@secure()
param applicationInsightsConnectionString string = ''

@description('Azure AD Tenant ID for Graph API authentication')
param azureTenantId string = ''

@description('Azure AD Client ID (App Registration) for Microsoft Graph API')
param azureClientId string = ''

@description('Azure AD Client Secret for Microsoft Graph API')
@secure()
param azureClientSecret string = ''

@description('SharePoint/OneDrive URL for the trade tracker document library')
param tradeTrackerUrl string = ''

@description('Excel file name for the trade tracker workbook')
param tradeTrackerFile string = ''

@description('Container image name with tag')
param imageName string = ''

@description('Enable Microsoft Entra ID authentication')
param enableEntraAuth bool = false

@description('Microsoft Entra ID Client ID (Application ID) - used as audience for token validation')
param entraClientId string = ''

@description('Microsoft Entra ID Tenant ID')
param entraTenantId string = ''

// Use placeholder image for initial deployment, azd will update with actual image
// imageName from azd already contains the full registry path (e.g., acr.azurecr.io/repo:tag)
var containerImage = !empty(imageName) ? imageName : 'mcr.microsoft.com/k8se/quickstart:latest'

// Reference existing ACR for credentials
resource containerRegistry 'Microsoft.ContainerRegistry/registries@2023-11-01-preview' existing = {
  name: containerRegistryName
}

resource containerApp 'Microsoft.App/containerApps@2024-03-01' = {
  name: name
  location: location
  tags: tags
  properties: {
    managedEnvironmentId: containerAppsEnvironmentId
    configuration: {
      activeRevisionsMode: 'Single'
      ingress: {
        external: true
        targetPort: 3000
        transport: 'http'
        allowInsecure: false
        traffic: [
          {
            weight: 100
            latestRevision: true
          }
        ]
        corsPolicy: {
          allowedOrigins: ['*']
          allowedMethods: ['GET', 'POST', 'OPTIONS']
          allowedHeaders: ['*']
          exposeHeaders: ['*']
          maxAge: 86400
        }
      }
      registries: [
        {
          server: containerRegistryLoginServer
          username: containerRegistry.listCredentials().username
          passwordSecretRef: 'acr-password'
        }
      ]
      secrets: [
        {
          name: 'acr-password'
          value: containerRegistry.listCredentials().passwords[0].value
        }
        {
          name: 'azure-tenant-id'
          value: azureTenantId
        }
        {
          name: 'azure-client-id'
          value: azureClientId
        }
        {
          name: 'azure-client-secret'
          value: azureClientSecret
        }
        {
          name: 'appinsights-connection-string'
          value: applicationInsightsConnectionString
        }
      ]
    }
    template: {
      containers: [
        {
          name: 'mcp-excel-server'
          image: containerImage
          resources: {
            cpu: json('0.5')
            memory: '1Gi'
          }
          env: [
            {
              name: 'PORT'
              value: '3000'
            }
            {
              name: 'AZURE_TENANT_ID'
              secretRef: 'azure-tenant-id'
            }
            {
              name: 'AZURE_CLIENT_ID'
              secretRef: 'azure-client-id'
            }
            {
              name: 'AZURE_CLIENT_SECRET'
              secretRef: 'azure-client-secret'
            }
            {
              name: 'APPLICATIONINSIGHTS_CONNECTION_STRING'
              secretRef: 'appinsights-connection-string'
            }
            {
              name: 'TRADE_TRACKER_URL'
              value: tradeTrackerUrl
            }
            {
              name: 'TRADE_TRACKER_FILE'
              value: tradeTrackerFile
            }
            {
              name: 'PYTHONUNBUFFERED'
              value: '1'
            }
            // Microsoft Entra ID authentication (token validation)
            // These env vars enable the server to validate incoming bearer tokens
            ...(enableEntraAuth && !empty(entraClientId) && !empty(entraTenantId) ? [
              {
                name: 'ENTRA_TENANT_ID'
                value: entraTenantId
              }
              {
                name: 'ENTRA_CLIENT_ID'
                value: entraClientId
              }
            ] : [])
          ]
          probes: [
            {
              type: 'Liveness'
              httpGet: {
                path: '/health'
                port: 3000
              }
              initialDelaySeconds: 10
              periodSeconds: 30
              failureThreshold: 3
            }
            {
              type: 'Readiness'
              httpGet: {
                path: '/health'
                port: 3000
              }
              initialDelaySeconds: 5
              periodSeconds: 10
              failureThreshold: 3
            }
          ]
        }
      ]
      scale: {
        minReplicas: 0
        maxReplicas: 10
        rules: [
          {
            name: 'http-scaling'
            http: {
              metadata: {
                concurrentRequests: '100'
              }
            }
          }
        ]
      }
    }
  }
}

// Note: Microsoft Entra ID authentication is handled via token validation in the application code
// When ENTRA_TENANT_ID and ENTRA_CLIENT_ID are set, the server validates incoming bearer tokens
// from AI Foundry's Project Managed Identity or Agent Identity

output id string = containerApp.id
output name string = containerApp.name
output fqdn string = containerApp.properties.configuration.ingress.fqdn
output authEnabled bool = enableEntraAuth
