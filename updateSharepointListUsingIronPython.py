# Importing all the necessary libraries

import System
from System import *
from System.IO import *
from System.Net import *
from System.Security import *
from System.Text import *
from System.Collections.Generic import Dictionary 
from System import Array,Guid,String,Object, DateTime
from Spotfire.Dxp.Data import *
from System.Net import ServicePointManager, SecurityProtocolType
from System.Collections.Generic import List
from Spotfire.Dxp.Data.Import import *
import json

# define the client ID, secret and token IDs
clientID = "<ClientID>"
clientSecret = "<clientSecret>
tenantID = <tenantID>

# Creat the API URL to get the access token
uriString = "https://accounts.accesscontrol.windows.net/" + tenantID + "/tokens/OAuth/2"  # authenticaiton URL
uri = Uri(uriString)

try:
	######  Get the access token #######

	webRequest = HttpWebRequest.Create(uri)

	# define the headers
	webRequest.ContentType =  "application/x-www-form-urlencoded"  # encoding the body in the URL as the tenant supports only form data or URL encoded data
	webRequest.Accept = "*/*"
	webRequest.Timeout = 1000000000
	webRequest.Method = "POST"

	# Body must contain - grant_type, resource, clientId and clientSecret
	form = "&grant_type=client_credentials&resource=00000003-0000-0ff1-ce00-000000000000/<sharepointSiteDomain>@" + tenantID + "&client_id=" + <clientId> + "@" + <tenantId> + "&client_secret=" + <clientSecret>

	"""
	the boy must contain the following parameters: grant_type, resource, client ID (<clientID>@<tenantID>), client Secret
	Resource is of the format:resource/siteDomain@tenantID
	1. resource: constant value - 00000003-0000-0ff1-ce00-000000000000
	2. SiteDomain: sharepoint site URL
	3. tenantID : pre-defined tenant ID
	"""

	# call the API
	encoding = Text.UTF8Encoding()
	byte1 = encoding.GetBytes(form)
	rqt = webRequest.GetRequestStream()
	rqt.Write(byte1,0,byte1.Length)

	rsp = webRequest.GetResponse()
	readStream = StreamReader(rsp.GetResponseStream(), Text.Encoding.UTF8)
	res = readStream.ReadToEnd()

	# convert the response in String format into json
	access_token = json.loads(res)

	# extract access_token key-value (storing this in a document property in the DXP file)
	#Document.Properties["pcAccessToken"] = access_token["access_token"]
	readStream.Close()
	rsp.Close()

except System.Net.WebException, we:
	print "Web Exception:", we.Message
	if we.Status == Net.WebExceptionStatus.ProtocolError: 
		print int(we.Response.StatusCode), we.Response.StatusDescription
		print (StreamReader(we.Response.GetResponseStream(), Text.Encoding.UTF8)).ReadToEnd()
except Exception, e:
	print e



uriString ="https://<domain.>sharepoint.com/sites/<site name>/_api/web/Lists/getbytitle('<listName>')/Items(<row number>)"  # sharepoint site->list->item URL is hardcoded in Document property for the purpose of this POC

type = "application/json;odata=verbose"

uri = Uri(uriString)

try:
	jsonOutput = '{"<column name>":"<new cell value>","__metadata": { "type": "SP.Data.<listname>ListItem"}}'  # body of the API call should contain the field that's being updated and metadata

	######  Update the Sharepoint List #######

	webRequest = HttpWebRequest.Create(uri)

	encoded = Document.Properties["pcAccessToken"]
	webRequest.Headers.Add("Authorization", "Bearer " + encoded) # send the ancoded access token
	
	webRequest.Headers.Add("X-HTTP-Method", "MERGE")  # this is the header which makes sure an existing item is updated
	webRequest.Headers.Add("If-Match", "*") # only if the item is found - in this case its item(1), which exists

	# setup other parameters
	webRequest.ContentType = type
	webRequest.Accept = type
	webRequest.Timeout = 1000000000
	webRequest.Method = "POST"

	# call the API
	encoding = Text.UTF8Encoding()
	byte1 = encoding.GetBytes(jsonOutput)

	rqt = webRequest.GetRequestStream()
	rqt.Write(byte1,0,byte1.Length)
	rsp = webRequest.GetResponse()
	readStream = StreamReader(rsp.GetResponseStream(), Text.Encoding.UTF8)
	res = readStream.ReadToEnd()

	readStream.Close()
	rsp.Close()

except System.Net.WebException, we:
	print "Web Exception:", we.Message
	if we.Status == Net.WebExceptionStatus.ProtocolError: 
		print int(we.Response.StatusCode), we.Response.StatusDescription
		print (StreamReader(we.Response.GetResponseStream(), Text.Encoding.UTF8)).ReadToEnd()
except Exception, e:
	print e
