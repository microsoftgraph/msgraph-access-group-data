// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Security;
using System.Threading.Tasks;
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

      var client = GetAuthenticatedGraphClient(config);

      var profileResponse = client.Me.Request().GetAsync().Result;
      Console.WriteLine("Hello " + profileResponse.DisplayName);

      // request 1 - all groups member of
      // Console.WriteLine("\n\nREQUEST 1 - ALL GROUPS MEMBER OF:");
      // var requestGroupsMemberOf = client.Me.MemberOf.Request();
      // var resultsGroupsMemberOf = requestGroupsMemberOf.GetAsync().Result;

      // foreach (var groupDirectoryObject in resultsGroupsMemberOf)
      // {
      //   var group = groupDirectoryObject as Microsoft.Graph.Group;
      //   var role = groupDirectoryObject as Microsoft.Graph.DirectoryRole;
      //   if (group != null)
      //   {
      //     Console.WriteLine("Group: " + group.Id + ": " + group.DisplayName);
      //   }
      //   else if (role != null)
      //   {
      //     Console.WriteLine("Role: " + role.Id + ": " + role.DisplayName);
      //   }
      //   else
      //   {
      //     Console.WriteLine(groupDirectoryObject.ODataType + ": " + groupDirectoryObject.Id);
      //   }
      // }

      // request 2 - all groups owner of
      Console.WriteLine("\n\nREQUEST 2 - ALL GROUPS OWNER OF:");
      var requestOwnerOf = client.Me.OwnedObjects.Request();
      var resultsOwnerOf = requestOwnerOf.GetAsync().Result;
      foreach (var ownedObject in resultsOwnerOf)
      {
        var group = ownedObject as Microsoft.Graph.Group;
        var role = ownedObject as Microsoft.Graph.DirectoryRole;
        if (group != null)
        {
          Console.WriteLine("Office 365 Group: " + group.Id + ": " + group.DisplayName);
        }
        else if (role != null)
        {
          Console.WriteLine("  Security Group: " + role.Id + ": " + role.DisplayName);
        }
        else
        {
          Console.WriteLine(ownedObject.ODataType + ": " + ownedObject.Id);
        }
      }

      Console.WriteLine("\nGraph Request:");
      Console.WriteLine(requestOwnerOf.GetHttpRequestMessage().RequestUri);
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

    private static IAuthenticationProvider CreateAuthorizationProvider(IConfigurationRoot config)
    {
      var clientId = config["applicationId"];
      var authority = $"https://login.microsoftonline.com/{config["tenantId"]}/v2.0";

      List<string> scopes = new List<string>();
      scopes.Add("https://graph.microsoft.com/.default");

      var cca = PublicClientApplicationBuilder.Create(clientId)
                                              .WithAuthority(authority)
                                              .WithDefaultRedirectUri()
                                              .Build();
      return MsalAuthenticationProvider.GetInstance(cca, scopes.ToArray());
    }

    private static GraphServiceClient GetAuthenticatedGraphClient(IConfigurationRoot config)
    {
      var authenticationProvider = CreateAuthorizationProvider(config);
      var graphClient = new GraphServiceClient(authenticationProvider);
      return graphClient;
    }
  }
}