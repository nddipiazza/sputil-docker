using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Security;
using System.Text.RegularExpressions;
using CommandLine;
using Microsoft.SharePoint.Client;
using System.Xml;
using System.Net.Http;
using System.Text;
using System.Web.Script.Serialization;

namespace SharepointTestUtility {

  // Commandline options
  class CmdOptions {
    [Option('d', "Domain", Required = true, HelpText = "Domain of sharepoint user")]
    public string Domain { get; set; }

    [Option('u', "Username", Required = true, HelpText = "Sharepoint Username. Include domain\\username if you are on prem")]
    public string Username { get; set; }

    [Option('p', "Password", Required = true, HelpText = "Sharepoint Password")]
    public string Password { get; set; }

    [Option('w', "WebApplicationUrl", Required = true, HelpText = "Sharepoint Web Application Url")]
    public string WebApplicationUrl { get; set; }

    [Option('P', "Port", Required = true, HelpText = "Sharepoint central admin console port")]
    public int AdminPort { get; set; }

    [Option('a', "ActionFile", Required = true, HelpText = "Json file describing what to do")]
    public string ActionFile { get; set; }
  }

  public class Util {
    public static string addSlashToUrlIfNeeded(string siteUrl) {
      string res = siteUrl;
      if (!res.EndsWith("/", StringComparison.CurrentCulture)) {
        res += "/";
      }
      return res;
    }
    public static string getBaseUrl(string siteUrl) {
      return new Uri(siteUrl).Scheme + "://" + new Uri(siteUrl).Host;
    }

    public static int getBaseUrlPort(string siteUrl) {
      return new Uri(siteUrl).Port;
    }

    public static string getBaseUrlHost(string siteUrl) {
      return new Uri(siteUrl).Host;
    }
    public static void deleteDirectory(string targetDir) {
      string[] files = Directory.GetFiles(targetDir);
      string[] dirs = Directory.GetDirectories(targetDir);

      foreach (string file in files) {
        System.IO.File.SetAttributes(file, FileAttributes.Normal);
        System.IO.File.Delete(file);
      }

      foreach (string dir in dirs) {
        deleteDirectory(dir);
      }

      Directory.Delete(targetDir, false);
    }

    public static bool isSharepointOnline(string url) {
      Regex rx = new Regex("https://[-a-zA-Z0-9]+\\.sharepoint\\.com",
          RegexOptions.Compiled | RegexOptions.IgnoreCase);

      return rx.IsMatch(url);
    }
  }

  public class Auth {
    public CredentialCache credentialsCache;
    public SharePointOnlineCredentials sharepointOnlineCredentials;
    public HttpClientHandler httpHandler;
    public Auth(string rootSite,
                bool isSharepointOnline,
                string domain,
                string username,
                string password,
                string authScheme) {
      if (!isSharepointOnline) {
        NetworkCredential networkCredential;
        if (password == null && username != null) {
          Console.WriteLine("Please enter password for {0}", username);
          networkCredential = new NetworkCredential(username, GetPassword(), domain);
        } else if (username != null) {
          networkCredential = new NetworkCredential(username, password, domain);
        } else {
          networkCredential = CredentialCache.DefaultNetworkCredentials;
        }
        credentialsCache = new CredentialCache();
        credentialsCache.Add(new Uri(rootSite), authScheme, networkCredential);
        CredentialCache credentialCache = new CredentialCache { { Util.getBaseUrlHost(rootSite), Util.getBaseUrlPort(rootSite), authScheme, networkCredential } };
        httpHandler = new HttpClientHandler() {
          CookieContainer = new CookieContainer(),
          Credentials = credentialCache.GetCredential(Util.getBaseUrlHost(rootSite), Util.getBaseUrlPort(rootSite), authScheme)
        };
      } else {
        SecureString securePassword = new SecureString();
        foreach (char c in password) {
          securePassword.AppendChar(c);
        }
        sharepointOnlineCredentials = new SharePointOnlineCredentials(username, securePassword);
      }

    }
    SecureString GetPassword() {
      var pwd = new SecureString();
      while (true) {
        ConsoleKeyInfo i = Console.ReadKey(true);
        if (i.Key == ConsoleKey.Enter) {
          break;
        }
        if (i.Key == ConsoleKey.Backspace) {
          if (pwd.Length > 0) {
            pwd.RemoveAt(pwd.Length - 1);
            Console.Write("\b \b");
          }
        } else {
          pwd.AppendChar(i.KeyChar);
          Console.Write("*");
        }
      }
      return pwd;
    }
  }

  class Program {

    Auth auth;
    string outputPath;
    string adminUrl;
    string username;
    string password;
    List<Dictionary<string, object>> actions;

    static SecureString GetSecureString(string input) {
      if (string.IsNullOrEmpty(input))
        throw new ArgumentException("Input string is empty and cannot be made into a SecureString", nameof(input));

      var secureString = new SecureString();
      foreach (char c in input)
        secureString.AppendChar(c);

      return secureString;
    }

    public ClientContext getClientContext(string site) {
      ClientContext clientContext = new ClientContext(site.Replace("https://", "http://"));
      clientContext.RequestTimeout = -1;
      if (auth.credentialsCache != null) {
        clientContext.Credentials = auth.credentialsCache;
      } else if (auth.sharepointOnlineCredentials != null) {
        clientContext.Credentials = auth.sharepointOnlineCredentials;
      }
      clientContext.ExecutingWebRequest += new EventHandler<WebRequestEventArgs>((obj, e) => {
        e.WebRequestExecutor.WebRequest.Headers.Add("X-FORMS_BASED_AUTH_ACCEPTED", "f");
      });
      return clientContext;
    }

    HttpWebRequest CreateSoapWebRequest(string action) {
      HttpWebRequest webRequest = (HttpWebRequest)WebRequest.Create(removeHttpsFromUrl(adminUrl) + "/_vti_adm/Admin.asmx");
      webRequest.Headers.Add("SOAPAction", "http://schemas.microsoft.com/sharepoint/soap/" + action);
      webRequest.ContentType = "text/xml; charset=\"utf-8\"";
      webRequest.Method = "POST";
      webRequest.Credentials = new NetworkCredential(username, password);
      return webRequest;
    }

    XmlDocument CreateSiteCollectionSoapEnvelope(string url, string title, string description, string user, string lcid, string webTemplate) {
      XmlDocument soapEnvelopeDocument = new XmlDocument();
      string soapEnv = string.Format(@"<?xml version=""1.0"" encoding=""utf-8""?>
        <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                       xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
          <soap:Body>
            <CreateSite xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">
              <Url>{0}</Url>
              <Title>{1}</Title>
              <Description>{2}</Description>
              <Lcid>{3}</Lcid>
              <WebTemplate>{4}</WebTemplate>
              <OwnerLogin>{5}</OwnerLogin>
              <OwnerName>{6}</OwnerName>
              <OwnerEmail/>
              <PortalUrl/>
              <PortalName/>
            </CreateSite>
          </soap:Body>
        </soap:Envelope>", url, title, description, lcid, webTemplate, user, user);
      Console.WriteLine("Create site collection soap request: {0}", soapEnv);
      soapEnvelopeDocument.LoadXml(soapEnv);
      return soapEnvelopeDocument;
    }

    XmlDocument DeleteSiteCollectionSoapEnvelope(string url) {
      XmlDocument soapEnvelopeDocument = new XmlDocument();
      string soapEnv = string.Format(@"<?xml version=""1.0"" encoding=""utf-8""?>
        <soap:Envelope xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/""
                       xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"">
          <soap:Body>
            <DeleteSite xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">
              <Url>{0}</Url>
            </DeleteSite>
          </soap:Body>
        </soap:Envelope>", removeHttpsFromUrl(url));
      Console.WriteLine("Delete site collection soap request: {0}", soapEnv);
      soapEnvelopeDocument.LoadXml(soapEnv);
      return soapEnvelopeDocument;
    }

    void InsertSoapEnvelopeIntoWebRequest(XmlDocument soapEnvelopeXml, HttpWebRequest webRequest) {
      using (Stream stream = webRequest.GetRequestStream()) {
        soapEnvelopeXml.Save(stream);
      }
    }

    public Program(CmdOptions options) {
      string json = System.IO.File.ReadAllText(options.ActionFile);
      actions = new JavaScriptSerializer().Deserialize<List<Dictionary<string, object>>>(json);
      outputPath = options.ActionFile + ".out";
      adminUrl = removeHttpsFromUrl(options.WebApplicationUrl) + ":" + options.AdminPort;
      auth = new Auth(removeHttpsFromUrl(options.WebApplicationUrl), Util.isSharepointOnline(options.WebApplicationUrl), options.Domain, options.Username, options.Password, "NTLM");
      username = options.Username;
      password = options.Password;

    }

    private String removeHttpsFromUrl(string url) {
      return url.Replace("https://", "http://");
    }

    private Folder EnsureFolder(ClientContext ctx, Folder ParentFolder, string FolderPath) {
      //Split up the incoming path so we have the first element as the a new sub-folder name 
      //and add it to ParentFolder folders collection
      string[] PathElements = FolderPath.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
      string Head = PathElements[0];
      Folder NewFolder = ParentFolder.Folders.Add(Head);
      ctx.Load(NewFolder);
      ctx.ExecuteQuery();

      //If we have subfolders to create then the length of PathElements will be greater than 1
      if (PathElements.Length > 1) {
        //If we have more nested folders to create then reassemble the folder path using what we have left i.e. the tail
        string Tail = string.Empty;
        for (int i = 1; i < PathElements.Length; i++)
          Tail = Tail + "/" + PathElements[i];

        //Then make a recursive call to create the next subfolder
        return EnsureFolder(ctx, NewFolder, Tail);
      } else
        //This ensures that the folder at the end of the chain gets returned
        return NewFolder;
    }

    //Overloaded main, called with CmdOptions from main(string[])
    private static void Main(CmdOptions options) {
      Program program = new Program(options);
      try {
        program.exec();
      } catch (WebException ex) {
        using (var reader = new System.IO.StreamReader(ex.Response.GetResponseStream())) {
            Console.WriteLine("Error: could not run {0}. Got a web exception with response {1}", program.actions, reader.ReadToEnd());
        }
        throw ex;
      } catch (Exception ex) {
        Console.WriteLine("Error: Could not run {0}. Exception: {1}", program.actions, ex);
        throw ex;
      } 
    }

    void exec() {
      List<string> res = new List<string>();
      foreach (Dictionary<string, object> action in actions) {
        string actionType = (string)action["Type"];
        if (actionType.Equals("createSiteCollection")) {
          XmlDocument soapEnvelopeXml = CreateSiteCollectionSoapEnvelope(
            removeHttpsFromUrl((string)action["Url"]),
            (string)action["Title"],
            (string)action["Description"],
            (string)action["User"],
            (string)action["Lcid"],
            (string)action["WebTemplate"]
          );
          HttpWebRequest webRequest = CreateSoapWebRequest("CreateSite");
          InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);
          IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);

          asyncResult.AsyncWaitHandle.WaitOne();

          using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult)) {
            using (StreamReader rd = new StreamReader(webResponse.GetResponseStream())) {
              Console.WriteLine(rd.ReadToEnd());
            }
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("deleteSiteCollection")) {
          XmlDocument soapEnvelopeXml = DeleteSiteCollectionSoapEnvelope(removeHttpsFromUrl((string)action["Url"]));
          HttpWebRequest webRequest = CreateSoapWebRequest("DeleteSite");
          InsertSoapEnvelopeIntoWebRequest(soapEnvelopeXml, webRequest);
          IAsyncResult asyncResult = webRequest.BeginGetResponse(null, null);

          asyncResult.AsyncWaitHandle.WaitOne();

          using (WebResponse webResponse = webRequest.EndGetResponse(asyncResult)) {
            using (StreamReader rd = new StreamReader(webResponse.GetResponseStream())) {
              Console.WriteLine(rd.ReadToEnd());
            }
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("createSite")) {
          WebCreationInformation webCreationInformation = new WebCreationInformation();
          webCreationInformation.WebTemplate = "STS#0";
          webCreationInformation.Description = (string)action["Description"];
          webCreationInformation.Title = (string)action["Title"];
          webCreationInformation.Url = removeHttpsFromUrl((string)action["Url"]);
          webCreationInformation.UseSamePermissionsAsParentSite = (bool)action["UseSamePermissionsAsParentSite"];
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            var site = clientContext.Web.Webs.Add(webCreationInformation);
            clientContext.Load(site);
            clientContext.ExecuteQuery();
            res.Add(site.Id.ToString());
            Console.WriteLine("Created site Guid={0}, Url={1}", site.Id, action["ParentSiteUrl"] + "/" + action["Url"]);
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("deleteSite")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["Url"]))) {
            clientContext.Web.DeleteObject();
            clientContext.ExecuteQuery();
            Console.WriteLine("Deleted site {0}", action["Url"]);
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("createList")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            Web web = clientContext.Web;
            clientContext.Load(web);
            clientContext.ExecuteQuery();

            //Create a List.
            ListCreationInformation listCreationInfo;
            List list;

            listCreationInfo = new ListCreationInformation();
            listCreationInfo.Description = (string)action["Description"];
            listCreationInfo.Title = (string)action["Title"];
            Enum.TryParse((string)action["ListTemplateName"], out ListTemplateType type);
            listCreationInfo.TemplateType = (int)type;

            list = web.Lists.Add(listCreationInfo);
            clientContext.ExecuteQuery();

            clientContext.Load(list, l => l.Id);
            clientContext.ExecuteQuery();

            res.Add(list.Id.ToString());

            Console.WriteLine("Created list - new GUID is {0}", list.Id);
          }
        } else if (actionType.Equals("deleteList")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            Web web = clientContext.Web;
            List oList = web.Lists.GetById(Guid.Parse((string)action["Guid"]));

            oList.DeleteObject();

            Console.WriteLine("Deleted list with GUID {0}", action["Guid"]);

            clientContext.ExecuteQuery();
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("createListItem")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();

            List spList = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

            ListItem oListItem = spList.AddItem(itemCreateInfo);

            Dictionary<string, object> columns = (Dictionary<string, object>)action["Columns"];

            foreach (string columnName in columns.Keys) {
              oListItem[columnName] = columns[columnName];
            }

            oListItem.Update();
            clientContext.Load(oListItem);
            clientContext.ExecuteQuery();

            Console.WriteLine("List item created id={0}", oListItem.Id);

            res.Add("" + oListItem.Id);
          }
        } else if (actionType.Equals("createFolder")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
            clientContext.Load(list, l => l.RootFolder.Name);
            clientContext.ExecuteQuery();
            ListItem newListItem = null;
            if (action.ContainsKey("ParentFolder")) {
              Folder parent = clientContext.Web.GetFolderByServerRelativeUrl(list.RootFolder.Name + "/" + (string)action["ParentFolder"]);
              Folder newFolder = EnsureFolder(clientContext, parent, (string)action["FolderName"]);
              Console.WriteLine("Sub folder created {0}", newFolder.Name);
              clientContext.ExecuteQuery();

              clientContext.Load(newFolder, f => f.ListItemAllFields.Id);
              clientContext.ExecuteQuery();

              Console.WriteLine("Sub folder created Name={0}, Id={1}", newFolder.Name, newFolder.ListItemAllFields.Id);

              res.Add("" + newFolder.ListItemAllFields.Id);
            } else {
              var folder = list.RootFolder;
              clientContext.Load(folder);
              clientContext.ExecuteQuery();

              ListItemCreationInformation newItemInfo = new ListItemCreationInformation();
              newItemInfo.UnderlyingObjectType = FileSystemObjectType.Folder;
              newItemInfo.LeafName = (string)action["FolderName"];
              newListItem = list.AddItem(newItemInfo);
              newListItem["Title"] = (string)action["FolderName"];
              newListItem.Update();
              clientContext.ExecuteQuery();
              clientContext.Load(newListItem, i => i.Id);
              clientContext.ExecuteQuery();
              Console.WriteLine("Folder created with id={0}", newListItem.Id);
              res.Add("" + newListItem.Id);
            }
          }
        } else if (actionType.Equals("deleteFolder")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {

            //List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
            //clientContext.Load(list, l => l.RootFolder.Name);
            //clientContext.ExecuteQuery();
            //Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(list.RootFolder.Name + "/" + (string)action["FolderUrl"]);

            List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
            clientContext.Load(list, l => l.RootFolder.Name);
            clientContext.ExecuteQuery();
            string folderUrl = (string)action["FolderUrl"];  // watch out with spaces!!
            string FolderToDeleteURL = list.RootFolder.Name + "/" + folderUrl;
            Web web = clientContext.Web;
            Folder folderToDelete = web.GetFolderByServerRelativeUrl(FolderToDeleteURL);
            clientContext.Load(folderToDelete);
            clientContext.ExecuteQuery();

            folderToDelete.DeleteObject();
            clientContext.ExecuteQuery();
            Console.WriteLine("Folder deleted with url={0}", action["FolderUrl"]);

            res.Add("");
          }
        } else if (actionType.Equals("createTextDocument")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            FileCreationInformation createFile = new FileCreationInformation();
            createFile.Url = (string)action["FileName"];
            //use byte array to set content of the file
            string somestring = (string)action["Text"];
            byte[] toBytes = Encoding.ASCII.GetBytes(somestring);

            createFile.Content = toBytes;

            List spList = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

            clientContext.Load(spList);
            clientContext.Load(spList.RootFolder);
            clientContext.Load(spList.RootFolder.Files);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.File addedFile;

            if (action.ContainsKey("ParentFolder")) {
              Folder folder = clientContext.Web.GetFolderByServerRelativeUrl(spList.RootFolder.Name + "/" + (string)action["ParentFolder"]);
              addedFile = folder.Files.Add(createFile);
            } else {
              addedFile = spList.RootFolder.Files.Add(createFile);
            }

            clientContext.Load(addedFile);
            clientContext.ExecuteQuery();

            clientContext.Load(addedFile, a => a.ListItemAllFields);
            clientContext.ExecuteQuery();

            ListItem item = addedFile.ListItemAllFields;
            item["Title"] = (string)action["Title"];
            item.Update();
            clientContext.Load(item);
            clientContext.ExecuteQuery();

            Console.WriteLine("Text file uploaded id={0}", item.Id);

            res.Add("" + item.Id);
          }
        } else if (actionType.Equals("deleteListItem")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
            ListItem oListItem = list.GetItemById((int)action["ItemId"]);
            oListItem.DeleteObject();
            clientContext.ExecuteQuery();

            Console.WriteLine("Deleted list item {0} from list guid {1}", action["ItemId"], action["ListGuid"]);
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("createListItemAttachment")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
            ListItem oListItem = list.GetItemById((int)action["ItemId"]);
            var attInfo = new AttachmentCreationInformation();
            attInfo.FileName = (string)action["FileName"];
            attInfo.ContentStream = new MemoryStream(System.IO.File.ReadAllBytes((string)action["FilePath"]));
            Attachment att = oListItem.AttachmentFiles.Add(attInfo);
            clientContext.Load(att);
            clientContext.ExecuteQuery();
            Console.WriteLine("Created list item attachment {0} on ListGuid={1}, ItemId={2}", action["FileName"], action["ListGuid"], action["ItemId"]);
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("deleteListItemAttachment")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
            ListItem oListItem = list.GetItemById((int)action["ItemId"]);
            Attachment att = oListItem.AttachmentFiles.GetByFileName((string)action["FileName"]);
            att.DeleteObject();
            clientContext.ExecuteQuery();
            Console.WriteLine("Deleted list item attachment {0} on ListGuid={1}, ItemId={2}", action["FileName"], action["ListGuid"], action["ItemId"]);
          }
          res.Add(""); // nothing to return
        } else if (actionType.Equals("createGroup")) { // Create a sharepoint group
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteCollectionUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            GroupCreationInformation groupCreationInformation = new GroupCreationInformation();
            groupCreationInformation.Title = (string)action["Title"];
            groupCreationInformation.Description = (string)action["Description"];

            Microsoft.SharePoint.Client.Group group = clientContext.Site.RootWeb.SiteGroups.Add(groupCreationInformation);

            clientContext.ExecuteQuery();

            if (action.ContainsKey("Role") && !"None".Equals(action["Role"])) {
              var role = new RoleDefinitionBindingCollection(clientContext);
              role.Add(clientContext.Site.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), (string)action["Role"])));
              clientContext.Site.RootWeb.RoleAssignments.Add(group, role);
              clientContext.ExecuteQuery();
            }

            clientContext.Load(group, g => g.Id);
            clientContext.ExecuteQuery();
            res.Add("" + group.Id);
          }
        } else if (actionType.Equals("deleteGroup")) { // Delete a sharepoint group
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteCollectionUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            if (action.ContainsKey("GroupId")) {
              clientContext.Site.RootWeb.SiteGroups.RemoveById((int)action["GroupId"]);
            } else {
              clientContext.Site.RootWeb.SiteGroups.RemoveByLoginName((string)action["GroupName"]);
            }

            clientContext.Site.RootWeb.Update();
            clientContext.ExecuteQuery();

            res.Add("");
          }
        } else if (actionType.Equals("createUser")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteCollectionUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            UserCreationInformation userCreationInfo = new UserCreationInformation();
            userCreationInfo.LoginName = (string)action["LoginName"]; // domain\username
            userCreationInfo.Title = (string)action["Title"]; // username
            if (action.ContainsKey("Email")) {
              userCreationInfo.Email = (string)action["Email"]; // username@domain.com
            }
            User spUser = clientContext.Site.RootWeb.SiteUsers.Add(userCreationInfo);

            clientContext.ExecuteQuery();

            if ((bool)action["IsSiteAdmin"]) {
              clientContext.ExecuteQuery();
              spUser.IsSiteAdmin = true;
              spUser.Update();

              clientContext.Load(spUser);
              clientContext.ExecuteQuery();
            }

            clientContext.Load(spUser, u => u.Id);
            clientContext.ExecuteQuery();
            res.Add("" + spUser.Id);
          }
        } else if (actionType.Equals("updateUser")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteCollectionUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            User spUser = clientContext.Site.RootWeb.EnsureUser((string)action["LoginName"]);

            clientContext.ExecuteQuery();
            spUser.IsSiteAdmin = (bool)action["IsSiteAdmin"];
            spUser.Update();

            clientContext.Load(spUser);
            clientContext.ExecuteQuery();

            res.Add("");
          }
        } else if (actionType.Equals("createRoleAssignment")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            // Note this can be a security group or a user.
            string pid;
            Principal principal;
            if (action.ContainsKey("LoginName")) {
              pid = (string)action["LoginName"];
              principal = clientContext.Site.RootWeb.EnsureUser(pid);
            } else {
              pid = (string)action["GroupName"];
              principal = clientContext.Site.RootWeb.SiteGroups.GetByName(pid);
            }

            string target = (string)action["Target"];

            if ("siteCollection".Equals(target)) {
              RoleDefinitionBindingCollection role = new RoleDefinitionBindingCollection(clientContext);
              role.Add(clientContext.Site.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), (string)action["Role"])));
              var resultingRole = clientContext.Site.RootWeb.RoleAssignments.Add(principal, role);
              clientContext.ExecuteQuery();

              Console.WriteLine("Added principal {0} to role {1} to site collection {2}", pid, role.GetType(), action["SiteUrl"]);

              clientContext.Load(resultingRole, r => r.PrincipalId);
              clientContext.ExecuteQuery();

              res.Add("" + resultingRole.PrincipalId);
            } else if ("site".Equals(target)) {
              RoleDefinitionBindingCollection role = new RoleDefinitionBindingCollection(clientContext);
              role.Add(clientContext.Site.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), (string)action["Role"])));
              var resultingRole = clientContext.Web.RoleAssignments.Add(principal, role);
              clientContext.ExecuteQuery();

              Console.WriteLine("Added principal {0} to role {1} to site {2}", pid, role.GetType(), action["SiteUrl"]);

              clientContext.Load(resultingRole, r => r.PrincipalId);
              clientContext.ExecuteQuery();

              res.Add("" + resultingRole.PrincipalId);
            } else if ("list".Equals(target)) {
              RoleDefinitionBindingCollection role = new RoleDefinitionBindingCollection(clientContext);
              role.Add(clientContext.Site.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), (string)action["Role"])));

              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              var resultingRole = list.RoleAssignments.Add(principal, role);
              clientContext.ExecuteQuery();

              Console.WriteLine("Added principal {0} to role {1} to list {2}", pid, role.GetType(), action["ListGuid"]);

              clientContext.Load(resultingRole, r => r.PrincipalId);
              clientContext.ExecuteQuery();

              res.Add("" + resultingRole.PrincipalId);
            } else if ("listItem".Equals(target)) {
              RoleDefinitionBindingCollection role = new RoleDefinitionBindingCollection(clientContext);
              role.Add(clientContext.Site.RootWeb.RoleDefinitions.GetByType((RoleType)Enum.Parse(typeof(RoleType), (string)action["Role"])));

              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              ListItem listItem = list.GetItemById((int)action["ItemId"]);

              var resultingRole = listItem.RoleAssignments.Add(principal, role);
              clientContext.ExecuteQuery();

              Console.WriteLine("Added principal {0} to role {1} to List {2}'s Item {3}", pid, role.GetType(), action["ListGuid"], action["ItemId"]);

              clientContext.Load(resultingRole, r => r.PrincipalId);
              clientContext.ExecuteQuery();

              res.Add("" + resultingRole.PrincipalId);
            } else {
              throw new Exception("Unsupported target " + target);
            }
          }
        } else if (actionType.Equals("breakRoleInheritance")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            string target = (string)action["Target"];

            if ("site".Equals(target)) {

              clientContext.Web.BreakRoleInheritance((bool)action["CopyRoleAssignments"], (bool)action["ClearSubScopes"]);

              clientContext.Load(clientContext.Web);
              clientContext.ExecuteQuery();

              Console.WriteLine("Broke role inheritence on site {0} successfully.", action["SiteUrl"]);

              res.Add("");
            } else if ("list".Equals(target)) {

              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              list.BreakRoleInheritance((bool)action["CopyRoleAssignments"], (bool)action["ClearSubScopes"]);

              clientContext.Load(list);
              clientContext.ExecuteQuery();

              Console.WriteLine("Broke role inheritence on list {0} successfully.", action["ListGuid"]);

              res.Add("");
            } else if ("listItem".Equals(target)) {

              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              ListItem listItem = list.GetItemById((int)action["ItemId"]);

              listItem.BreakRoleInheritance((bool)action["CopyRoleAssignments"], (bool)action["ClearSubScopes"]);

              clientContext.Load(listItem);
              clientContext.ExecuteQuery();

              Console.WriteLine("Broke role inheritence on list={0}, listItem={1} successfully.", action["ListGuid"], action["ItemId"]);

              res.Add("");
            } else {
              throw new Exception("Unsupported target " + target);
            }
          }
        } else if (actionType.Equals("resetRoleInheritance")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            string target = (string)action["Target"];

            if ("site".Equals(target)) {

              clientContext.Web.ResetRoleInheritance();

              clientContext.Load(clientContext.Web);
              clientContext.ExecuteQuery();

              Console.WriteLine("Reset role inheritence on web {0} successfully.", action["SiteUrl"]);

              res.Add("");

            } else if ("list".Equals(target)) {
              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              list.ResetRoleInheritance();

              clientContext.Load(list);
              clientContext.ExecuteQuery();

              Console.WriteLine("Reset role inheritence on list {0} successfully.", action["ListGuid"]);

              res.Add("");

            } else if ("listItem".Equals(target)) {
              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              ListItem listItem = list.GetItemById((int)action["ItemId"]);

              listItem.ResetRoleInheritance();

              clientContext.Load(listItem);
              clientContext.ExecuteQuery();

              Console.WriteLine("Reset role inheritence on list {0}'s item {1} successfully.", action["ListGuid"], action["ItemId"]);

              res.Add("");

            } else {
              throw new Exception("Unsupported target " + target);
            }
          }
        } else if (actionType.Equals("deleteRoleAssignment")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            string target = (string)action["Target"];

            RoleAssignment resultingRole;
            if ("siteCollection".Equals(target)) {
              resultingRole = clientContext.Site.RootWeb.RoleAssignments.GetByPrincipalId((int)action["PrincipalId"]);
            } else if ("list".Equals(target)) {
              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
              resultingRole = list.RoleAssignments.GetByPrincipalId((int)action["PrincipalId"]);
            } else if ("listItem".Equals(target)) {
              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));
              ListItem listItem = list.GetItemById((int)action["ItemId"]);
              resultingRole = listItem.RoleAssignments.GetByPrincipalId((int)action["PrincipalId"]);
            } else {
              throw new Exception("Unsupported target " + target);
            }

            resultingRole.DeleteObject();
            clientContext.ExecuteQuery();
            Console.WriteLine("Deleted role assignment {0}", action["PrincipalId"]);
            res.Add("");
          }
        } else if (actionType.Equals("addUserToFolderPermissions")) {
          Console.WriteLine("SiteURL={0}", (string)action["ParentSiteUrl"]);
          Console.WriteLine("ListGuid={0}", (string)action["ListGuid"]);
          Console.WriteLine("FolderID={0}", (Int32.Parse(action["FolderID"].ToString())).ToString());
          Console.WriteLine("USer={0}", (string)action["user"]);


          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["ParentSiteUrl"]))) {
            Site site = clientContext.Site;
            Web web = clientContext.Web;
            List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

            // Enter server relative URL of the folder(int)action["FolderID"]
            int folderId = Int32.Parse(action["FolderID"].ToString());
            Console.WriteLine("Folder Id is {0}", folderId.ToString());
            ListItem FolderItem = list.GetItemById(folderId);
            clientContext.ExecuteQuery();

            //Folder newFolder    = list.RootFolder.Folders.Add("F4");
            //clientContext.ExecuteQuery();
            FolderItem.BreakRoleInheritance(false, true);
            var role = new RoleDefinitionBindingCollection(clientContext);
            role.Add(web.RoleDefinitions.GetByType(RoleType.Contributor));
            User user = web.EnsureUser((string)action["user"]);
            FolderItem.RoleAssignments.Add(user, role);
            FolderItem.Update();
            clientContext.ExecuteQuery();
            Console.WriteLine("Added to the premissions on folder - User Id is {0}", (string)action["user"]);
            Console.WriteLine("Folder Id is {0}", (string)action["FolderID"]);
            res.Add("Used added to permissions");

          }

        } else if (actionType.Equals("addUserToSharepointGroup")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteCollectionUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            User spUser = clientContext.Site.RootWeb.EnsureUser((string)action["LoginName"]);

            clientContext.Load(spUser);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.GroupCollection collGroup = clientContext.Web.SiteGroups;
            Microsoft.SharePoint.Client.Group oGroup;
            if (action.ContainsKey("GroupName")) {
              oGroup = collGroup.GetByName((string)action["GroupName"]);
            } else {
              oGroup = collGroup.GetById((int)action["GroupId"]);
            }

            oGroup.Users.AddUser(spUser);
            clientContext.ExecuteQuery();

            clientContext.Load(oGroup);
            clientContext.ExecuteQuery();

            Console.WriteLine("Added user {0} to sharepoint group {1}", spUser.LoginName, oGroup.LoginName);

            res.Add(spUser.LoginName);
          }
        } else if (actionType.Equals("removeUserFromSharepointGroup")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteCollectionUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            Microsoft.SharePoint.Client.GroupCollection collGroup = clientContext.Web.SiteGroups;
            Microsoft.SharePoint.Client.Group oGroup;
            if (action.ContainsKey("GroupName")) {
              oGroup = collGroup.GetByName((string)action["GroupName"]);
            } else {
              oGroup = collGroup.GetById((int)action["GroupId"]);
            }

            oGroup.Users.RemoveByLoginName((string)action["LoginName"]);
            clientContext.ExecuteQuery();

            res.Add("");
          }
        } else if (actionType.Equals("setNoIndex")) {
          using (ClientContext clientContext = getClientContext(removeHttpsFromUrl((string)action["SiteUrl"]))) {
            clientContext.Load(clientContext.Web);
            clientContext.Load(clientContext.Site);
            clientContext.Load(clientContext.Site.RootWeb);
            clientContext.ExecuteQuery();

            string target = (string)action["Target"];

            if ("site".Equals(target)) {
              clientContext.Load(clientContext.Web);
              clientContext.Load(clientContext.Site);
              clientContext.Load(clientContext.Site.RootWeb);
              clientContext.ExecuteQuery();

              clientContext.Web.NoCrawl = (bool)action["Value"];

              clientContext.Web.Update();
              clientContext.ExecuteQuery();

              res.Add("");
            } else if ("list".Equals(target)) {
              clientContext.Load(clientContext.Web);
              clientContext.Load(clientContext.Site);
              clientContext.Load(clientContext.Site.RootWeb);
              clientContext.ExecuteQuery();

              List list = clientContext.Web.Lists.GetById(Guid.Parse((string)action["ListGuid"]));

              list.NoCrawl = (bool)action["Value"];

              list.Update();

              clientContext.ExecuteQuery();

              res.Add("");
            } else {
              throw new Exception("Unsupported target " + target);
            }

            Console.WriteLine("Successfully set NoIndex = {0} to {1}", action["Value"], action["Target"]);
          }
        } else {
          throw new Exception("Unsupported action " + action);
        }

        // TODO cases that need added in order of priority

        // move list item
        // rename list item
        // create a mysite site collection
        // add field to list
        // remove field from list
      }
      System.IO.File.WriteAllLines(outputPath, res);
    }

    static void Main(string[] args) {
      var cmdOptions = Parser.Default.ParseArguments<CmdOptions>(args);
      cmdOptions.WithParsed(
          options => {
            Main(options);
          });
    }
  }
}
