from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from office365.sharepoint.files.file import File import os
#pip install SharePlum==0.5.0
#pip install Office365-REST-Python-Client
def UploadFile(arg):
  
  SPO_Url = arg['SPO_Url']    #https://siteaddress/sites/SiteName/
  SPO_ClientID = arg['SPO_ClientID']
  SPO_ClientSecret = arg['SPO_ClientSecret']
  SPO_FilePath = arg['SPO_FilePath']  #/sites/Name/filepath
  Local_Path = arg['LOCALPATH']   #localpath with "//""
  FileName = arg['File_Name']   #filename with extension
  
  app_settings = { 
            'url': SPO_Url,
            'client_id': SPO_ClientID,
            'client_secret': SPO_ClientSecret,
         
        }
  
  context_auth = AuthenticationContext(url=app_settings['url'])
  context_auth.acquire_token_for_app(client_id=app_settings['client_id'], client_secret=app_settings['client_secret'])
  
  ctx = ClientContext(SPO_Url, context_auth)
  web = ctx.web
  ctx.load(web)
  ctx.execute_query()
  print("Web site title: {0}".format(web.properties['Title']))
  
  path = Local_Path+FileName      
  with open(path, 'rb') as content_file:
    file_content = content_file.read()  
  target_url = SPO_FilePath+FileName.format(os.path.basename(path))
  File.save_binary(ctx, target_url, file_content)
  
  return 'successfully completed'