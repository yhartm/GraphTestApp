using System.Net.Http.Headers;

namespace GraphTutorial;

public class RequestBuilder
{
    private string _accessToken;
    private const string BaseUrl = "https://graph.microsoft.com/v1.0/me";

    public RequestBuilder(string accessToken)
    {
        _accessToken = accessToken;
    }

    public HttpRequestMessage GetRequest(string suffix = "")
    {
        var request = new HttpRequestMessage
        {
            RequestUri = new Uri(BaseUrl + suffix),
            Method = HttpMethod.Get
        };
        
        request.Headers.Authorization = new AuthenticationHeaderValue("bearer", _accessToken);

        return request;
    }

    public void RefreshToken(string token)
    {
        _accessToken = token;
    }
    
}