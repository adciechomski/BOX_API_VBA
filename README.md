# BOX_API_VBA
BOX authentication and BOX file upload example using VBA



This module presents utilization of BOXAuth module and BOXFileUpload module and how it can be used to interact with BOX doing HTTP calls using VBA.
BOXAuth module is crucial to do all later calls, whereas BOXFileUpload module is example of POST API call.

Please, remmember to add appropierte references to your project: ScriptingRuntime, Microsoft HTTP Object Library, Microsoft Internet Controls

Authentication function:
GetBoxAuthToken() 
returns dictionary with following keys assigned to token string parameters:
- access_token
- expires_in
- restricted_to
- refresh_token
- token_type
- ConnectionStatusBOX
