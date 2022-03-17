using System;
using System.Collections.Generic;
using System.Security;
using Microsoft.Identity.Client;
using Microsoft.Graph;
using Microsoft.Extensions.Configuration;
using Helpers;

namespace graphconsoleapp
{
  public class Program
  {
    public static void Main(string[] args)
    {
      var config = LoadAppSettings();
      if (config == null)
      {
        Console.WriteLine("Invalid appsettings.json file.");
        return;
      }
      var userName = ReadUsername();
      var userPassword = ReadPassword();

      var client = GetAuthenticatedGraphClient(config, userName, userPassword);
      var requestAllUsers = client.Users.Request();

      // all users
      // var results = requestAllUsers.GetAsync().Result;
      // foreach (var user in results)
      // {
      //   Console.WriteLine(user.Id + ": " + user.DisplayName + " <" + user.Mail + ">");
      // }

      // Console.WriteLine("\nGraph Request:");
      // Console.WriteLine(requestAllUsers.GetHttpRequestMessage().RequestUri);

      // request 1 - get me
      var requestMeUser = client.Me.Request();

      var resultMe = requestMeUser.GetAsync().Result;
      Console.WriteLine(resultMe.Id + ": " + resultMe.DisplayName + " <" + resultMe.Mail + ">");

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestMeUser.GetHttpRequestMessage().RequestUri);

      // request 1 - current user's photo

      var requestUserPhoto = client.Me.Photo.Request();
      var resultsUserPhoto = requestUserPhoto.GetAsync().Result;
      // display photo metadata
      Console.WriteLine("                Id: " + resultsUserPhoto.Id);
      Console.WriteLine("media content type: " + resultsUserPhoto.AdditionalData["@odata.mediaContentType"]);
      Console.WriteLine("        media etag: " + resultsUserPhoto.AdditionalData["@odata.mediaEtag"]);

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestUserPhoto.GetHttpRequestMessage().RequestUri);

      var requestUserPhotoFile = client.Me.Photo.Content.Request();
      var resultUserPhotoFile = requestUserPhotoFile.GetAsync().Result;

      var profilePhotoPath = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory(), "profilePhoto_" + resultsUserPhoto.Id + ".jpg");
      var profilePhotoFile = System.IO.File.Create(profilePhotoPath);
      resultUserPhotoFile.Seek(0, System.IO.SeekOrigin.Begin);
      resultUserPhotoFile.CopyTo(profilePhotoFile);
      Console.WriteLine("Saved file to: " + profilePhotoPath);

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestUserPhoto.GetHttpRequestMessage().RequestUri);

      // request 2 - user's manager
      var userId = "ca554b52-ccd3-46b5-84e4-7f6e2148fe35"; // bad data for testing
      var requestUserManager = client.Users[userId]
                                      .Manager
                                      .Request();
      var resultsUserManager = requestUserManager.GetAsync().Result;
      Console.WriteLine("   User: " + userId);
      Console.WriteLine("Manager: " + resultsUserManager.Id);
      var resultsUserManagerUser = resultsUserManager as Microsoft.Graph.User;
      if (resultsUserManagerUser != null)
      {
        Console.WriteLine("Manager: " + resultsUserManagerUser.DisplayName);
        Console.WriteLine(resultsUserManager.Id + ": " + resultsUserManagerUser.DisplayName + " <" + resultsUserManagerUser.Mail + ">");
      }

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestUserManager.GetHttpRequestMessage().RequestUri);
    }
    private static IConfigurationRoot? LoadAppSettings()
    {
      try
      {
        var config = new ConfigurationBuilder()
                          .SetBasePath(System.IO.Directory.GetCurrentDirectory())
                          .AddJsonFile("appsettings.json", false, true)
                          .Build();

        if (string.IsNullOrEmpty(config["applicationId"]) ||
            string.IsNullOrEmpty(config["tenantId"]))
        {
          return null;
        }

        return config;
      }
      catch (System.IO.FileNotFoundException)
      {
        return null;
      }
    }
    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var clientId = config["applicationId"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("User.Read");
      scopes.Add("User.Read.All");

      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray(), userName, userPassword);
    }
    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config, string userName, SecureString userPassword)
    {
      var authenticationProvider = CreateAuthorizationProvider(config, userName, userPassword);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }
    private static SecureString ReadPassword()
    {
      Console.WriteLine("Enter your password");
      SecureString password = new SecureString();
      while (true)
      {
        ConsoleKeyInfo c = Console.ReadKey(true);
        if (c.Key == ConsoleKey.Enter)
        {
          break;
        }
        password.AppendChar(c.KeyChar);
        Console.Write("*");
      }
      Console.WriteLine();
      return password;
    }
    private static string ReadUsername()
    {
      string? username;
      Console.WriteLine("Enter your username");
      username = Console.ReadLine();
      return username ?? "";
    }
  }
}