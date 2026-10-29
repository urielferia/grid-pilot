using System.Net;
namespace GridPilot.Tests;
public sealed class FailingHttp : HttpMessageHandler
{
    public int Calls {get;private set;}
    protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request,CancellationToken token){Calls++;return Task.FromResult(new HttpResponseMessage(HttpStatusCode.TooManyRequests){Content=new StringContent("No service called")});}
}
