param prefix string = 'test'
param location string = resourceGroup().location
param office365ConnectionName string = 'o365'

resource office365Connection 'Microsoft.Web/connections@2016-06-01' = {
  name: office365ConnectionName
  location: location
  properties: {
    displayName: office365ConnectionName
    api: {
      id: '${subscription().id}/providers/Microsoft.Web/locations/${location}/managedApis/office365'
    }
    parameterValues: {
      // Include necessary parameters for Office 365 connection
      // This often includes client ID, secret, and tenant ID for OAuth-based connections
      // Ensure these values are securely handled
    }
  }
}

resource logicApp 'Microsoft.Logic/workflows@2019-05-01' = {
  name: '${prefix}-logic-app'
  location: location
  properties: {
    state: 'Enabled'
    definition: {
      '$schema': 'https://schema.management.azure.com/providers/Microsoft.Logic/schemas/2016-06-01/workflowdefinition.json#'
      contentVersion: '1.0.0.0'
      parameters: {
        '$connections': {
          defaultValue: {}
          type: 'Object'
        }
      }
      triggers: {
        manual: {
          type: 'Request'
          kind: 'Http'
          inputs: {
            schema: {
              properties: {
                entityName: {
                  type: 'string'
                }
                entityType: {
                  type: 'string'
                }
                from: {
                  type: 'string'
                }
                link: {
                  type: 'string'
                }
                to: {
                  type: 'string'
                }
              }
              type: 'object'
            }
          }
        }
      }
      actions: {
        Send_an_email_V2: {
          type: 'ApiConnection'
          runAfter: {}
          inputs: {
            body: {
              Body: '<p>User @{triggerBody().from} would like to access @{triggerBody().entityName} @{triggerBody().entityType} content. You can grant permission on <a href="@{triggerBody().link}">vault permission page</a>.</p>'
              Importance: 'Normal'
              Subject: 'Key Vault Access Request'
              To: '@triggerBody().to'
            }
            host: {
              connection: {
                name: '@parameters(\'$connections\')[\'office365\'][\'connectionId\']'
              }
            }
            method: 'post'
            path: '/v2/Mail'
          }
        }
      }
      outputs: {}
    }
    parameters: {
      '$connections': {
        value: {
          office365: {
            connectionId: office365Connection.id
            connectionName: office365Connection.name
            id: '${subscription().id}/providers/Microsoft.Web/locations/${location}/managedApis/office365'
          }
        }
      }
    }
  }
}
