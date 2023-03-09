

using System.Net;
using Microsoft.Kiota.Abstractions;
using Microsoft.Kiota.Abstractions.Authentication;

public class AuthProvider : IAuthenticationProvider
{
    public string Token
    {
        get;
        set;
    }
    private string[] _scopes;
    private const string AuthorizationHeaderKey = "Authorization";
    private const string ClaimsKey = "claims";

    public AuthProvider(string token, string[] scopes)
    {
        Token = token;
        _scopes = scopes;
    }
    

    public async Task AuthenticateRequestAsync(RequestInformation request, Dictionary<string, object>? additionalAuthenticationContext = null,
        CancellationToken cancellationToken = new CancellationToken())
    {
        if(request == null) throw new ArgumentNullException(nameof(request));
        if(additionalAuthenticationContext != null &&
           additionalAuthenticationContext.ContainsKey(ClaimsKey) &&
           request.Headers.ContainsKey(AuthorizationHeaderKey))
            request.Headers.Remove(AuthorizationHeaderKey);

        if(!request.Headers.ContainsKey(AuthorizationHeaderKey))
        {
            var token = Token;
            if(!string.IsNullOrEmpty(token))
                request.Headers.Add(AuthorizationHeaderKey, $"Bearer {token}");
        }
    }
}