Option Explicit On

Imports Microsoft.Graph 'Add reference to: Graph 4.46.0
Imports Microsoft.Identity.Client

Module Module1

    Sub Main()
        'Addt'l Reference: https://learn.microsoft.com/en-us/graph/api/reportroot-getmailboxusagestorage?view=graph-rest-1.0


        ' Set the client ID, tenant ID, and client secret for your Azure AD app
        Dim ClientId As String = "-- YOUR CLIENT ID HERE --"
        Dim TenantId As String = "-- YOUR TENANT ID HERE --"
        Dim ClientSecret As String = "-- YOUR CLIENT SECRET VALUE HERE --"

        ' Set the scope for the Graph API call
        Dim scope As String = "https://graph.microsoft.com/.default"

        ' Create a new instance of the Confidential Client Application using the client ID, tenant ID, and client secret
        Dim confidentialClientApplication As IConfidentialClientApplication = ConfidentialClientApplicationBuilder _
            .Create(ClientId) _
            .WithTenantId(TenantId) _
            .WithClientSecret(ClientSecret) _
            .Build()

        ' Authenticate the application and retrieve an access token
        Dim authResult As AuthenticationResult = confidentialClientApplication _
            .AcquireTokenForClient(New String() {scope}) _
            .ExecuteAsync() _
            .GetAwaiter() _
            .GetResult()


        ' Create a new instance of the GraphServiceClient using the access token
        Dim graphClient As GraphServiceClient = New GraphServiceClient(New DelegateAuthenticationProvider(Function(requestMessage)
                                                                                                              requestMessage.Headers.Authorization = New System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", authResult.AccessToken)
                                                                                                              Return Task.FromResult(0)
                                                                                                          End Function))

        Dim req = graphClient.Reports.GetMailboxUsageDetail("D7").Request().GetHttpRequestMessage()
        Dim resp = graphClient.HttpProvider.SendAsync(req).Result

        Dim strReport As String = ""
        If resp.IsSuccessStatusCode Then
            strReport = resp.Content.ReadAsStringAsync().Result
            Console.WriteLine(strReport)  'this returns CSV data
        Else
            Console.WriteLine(resp.ReasonPhrase)
        End If

        Console.ReadLine()
    End Sub

End Module
