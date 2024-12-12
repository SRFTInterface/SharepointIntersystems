/// All API calls are made from here to SPO
/// Files are downloaded as Ens.Streamconainer's, everything is is responded as a EnsLib.HTTP.GenericMessage
Class Component.SharePointOnline.Operations.SharePointOnlineRESTOperation Extends (EnsLib.REST.Operation, %CSP.REST)
{

Parameter INVOCATION = "Queue";

/// The name of Site created within your tennacy you wish to interact with.<br/>
/// e.g. https://intersystems.sharepoint.com/sites/Test-Site<br/>
/// This will be URLEncoded automatically at the point of transmit
Property SharepointSitename As %String [ InitialExpression = "Test-Site", Required ];

/// The name of Document folder created within your Site you wish to interact with.<br /> 
/// e.g. https://intersystems.sharepoint.com/sites/Test-Site/Shared Documents/ <br /> 
/// This will be URLEncoded automatically at the point of transmit <br />
Property SharepointFilePath As %String [ InitialExpression = "Shared Documents", Required ];

/// Overwrite file if exists already on SPO site
Property OverwriteExistingFile As %String(VALUELIST = ",true,false") [ InitialExpression = "true", Required ];

/// If false gets all files in folder, if true it will get all files from last time a file was downloaded via this operation
Property GetOnlyNewFiles As %String(VALUELIST = ",true,false") [ InitialExpression = "true", Required ];

/// Default is don't need all the info off the response so just return the Status line 
Property CondensedHTMLResponse As %Boolean [ InitialExpression = 1, Required ];

/// The OAuth server application setup in TIE for downloading a Client Credential
Property OAuthClientApplicationName As %String(MAXLEN = 64, MINLEN = 1) [ InitialExpression = "SPO", Required ];

/// The scope you setup on your SPO Application in Azure
Property OAuthScope As %String(MAXLEN = 1024);

/// A unique name to track your access token
Property OAuthTokenSessionId As %String [ InitialExpression = "SPO", Required ];

/// For SPO the Params should be {"resource":"00000003-0000-0ff1-ce00-000000000000/{TennantName}.sharepoint.com@{TennantID}"} <p>
/// 00000003-0000-0ff1-ce00-000000000000 is the ResourceID asigned to SPO by Microsoft, and should not change <p>
/// {TennantName} should be changed to the same as the HTTP server e.g.intersystems.sharepoint.com <p>
/// {TennantID} is the ID for your server name tennancy 
Property Params As %String(MAXLEN = 500000) [ InitialExpression = "{'resource':'00000003-0000-0ff1-ce00-000000000000/{TennantName}.sharepoint.com@{TennantID}'}", Required ];

Parameter SETTINGS = "SharepointSitename:Basic,SharepointFilePath:Basic,OverwriteExistingFile:Basic,GetOnlyNewFiles:Basic,CondensedHTMLResponse:Basic,OAuthClientApplicationName:OAuth,OAuthScope:OAuth,OAuthTokenSessionId:OAuth, Params:OAuth";

ClassMethod OnGetConnections(Output pArray As %String, pItem As Ens.Config.Item)
{
	d pItem.PopulateModifiedSettings()
	s (tValue,tIndex)=""
	for
	{
		s tSetting = pItem.ModifiedSettings.GetNext(.tIndex) q:tIndex=""
		if tSetting.Name = "SharepointSitename" s tValue=tSetting.Value  q:tIndex=""
		If tSetting.Name = "SharepointFilePath" Set tValue=tValue_","_tSetting.Value q:tIndex=""
        if tSetting.Name = "OverwriteExistingFile" s tValue=tSetting.Value  q:tIndex=""
		If tSetting.Name = "GetOnlyNewFiles" Set tValue=tValue_","_tSetting.Value q:tIndex=""
        if tSetting.Name = "CondensedHTMLResponse" s tValue=tSetting.Value  q:tIndex=""
		If tSetting.Name = "OAuthClientApplicationName" Set tValue=tValue_","_tSetting.Value q:tIndex=""
        if tSetting.Name = "OAuthScope" s tValue=tSetting.Value  q:tIndex=""
		If tSetting.Name = "OAuthTokenSessionId" Set tValue=tValue_","_tSetting.Value q:tIndex=""
        If tSetting.Name = "Params" Set tValue=tValue_","_tSetting.Value q:tIndex=""
	}
	for i=1:1:$Length(tValue,",") 
	{
		s tOne=$ZStrip($p(tValue,",",i),"<>W")  Continue:""=tOne  s pArray(tOne)=""
	}
}

Method GetFileList(pQuery As Ens.Request, Output pResponse As EnsLib.HTTP.GenericMessage) As %Status
{
    $$$TRACE("GetFileList = "_..GetOnlyNewFiles_", id = "_$SYSTEM.SYS.ProcessID())
    s filter=""
    if (""_..GetOnlyNewFiles = "true")
    {
        s a="",b=..%RequestHeader.TargetConfigName
  
        $$$TRACE(b)
       //logical mode so status is 9
       //&SQL(SELECT Max(TimeProcessed) INTO :a FROM Ens.MessageHeader where TargetConfigName=:b  AND Status = 9 )
       &SQL(SELECT TOP 1 TimeProcessed INTO :a FROM Ens.MessageHeader head WHERE head.Status = 9 AND (head.SourceConfigName = :b AND head.MessageBodyClassName = 'Ens.StreamContainer') ORDER BY head.ID Desc)
           IF SQLCODE<0 {WRITE "SQLCODE error ",SQLCODE  RETURN %msg}
        //$$$Trace(a)
        //needs encoded to time for API as XSD time, a should be timestamp value hopefully
        set XSDtime=##Class(%Library.StringTimeStamp).LogicalToXSD(a) 
        
        $$$TRACE(XSDtime)
        //there's a Z on XSD time... 
        set XSDtime=$EXTRACT(XSDtime,1,$LENGTH(XSDtime)-1)
        $$$TRACE(XSDtime)
        s filter = "?$filter=TimeLastModified gt datetime'"_XSDtime_"'"
        s filtervalue = "TimeLastModified gt datetime'"_XSDtime_"'"
        s filter = "?$filter="_$$$URLENCODE("TimeLastModified gt datetime'"_XSDtime_"'")
        $$$TRACE("Filter = "_filter)

        $$$TRACE("SET FILTER AS hardcoded as it is working in postman with no spaces")
        //S filter="$filter"
        //s filter=$$$URLENCODE(filter)
       // $$$TRACE("URL ENCODED FILTER" _filter)
    }
    
    Set tHttpRequest=##class(%Net.HttpRequest).%New()
    set tHttpRequest.WriteRawMode=1
    do tHttpRequest.SetHeader("Accept","application/json;odata=verbose")
    do tHttpRequest.SetHeader("Content-Type","application/json")
    //if filter '=""{Do tHttpRequest.SetParam("$filter", filtervalue)}
    Set ..Adapter.URL="/sites/"_$$$URLENCODE(..SharepointSitename)_"/_api/web/GetFolderByServerRelativeUrl('"_$$$URLENCODE(..SharepointFilePath)_"')/Files"_filter
   //Set ..Adapter.URL="/sites/"_$$$URLENCODE(..SharepointSitename)_"/_api/web/GetFolderByServerRelativeUrl('"_$$$URLENCODE(..SharepointFilePath)_"')/Files$filter"
    $$$TRACE("Get changes from "_..Adapter.URL)
    Set tSC=..AddAccessToken(.tHttpRequest)
   //s tHttpRequest.EntityBody=##Class(%Stream.FileBinary).%New()
	//s sc=tHttpRequest.EntityBody.Write(pQuery.StringValue)
	Set tHttpResponse = ##class(%Net.HttpResponse).%New()
	S send="GET"
    $$$TRACE("Sending to "_..Adapter.URL)
	s tSC = ..SendRequest(.tHttpResponse,send,tHttpRequest, .pResponse)
    Quit ..constructResponse(.tHttpResponse,.pResponse)
}

Method DeleteFile(pRequest As Messages.Generic.GenericDocument, Output pResponse As EnsLib.HTTP.GenericMessage) As %Status
{
    $$$TRACE("Delete file "_pRequest.FileName)
    Set tHttpRequest=##class(%Net.HttpRequest).%New()
    set tHttpRequest.WriteRawMode=1
    do tHttpRequest.SetHeader("Accept","application/json;odata=verbose")
    do tHttpRequest.SetHeader("Content-Type","application/json")
    Set ..Adapter.URL="/sites/"_$$$URLENCODE(..SharepointSitename)_"/_api/web/GetFolderByServerRelativeUrl('"_$$$URLENCODE(..SharepointFilePath)_"')/Files('"_$$$URLENCODE(pRequest.FileName)_"')"
    $$$TRACE("Get changes from "_..Adapter.URL)
    Set tSC=..AddAccessToken(.tHttpRequest)
   //s tHttpRequest.EntityBody=##Class(%Stream.FileBinary).%New()
	//s sc=tHttpRequest.EntityBody.Write(pQuery.StringValue)
	Set tHttpResponse = ##class(%Net.HttpResponse).%New()
	S send="DELETE"
    $$$TRACE("Sending to "_..Adapter.URL)
	s tSC = ..SendRequest(.tHttpResponse,send,tHttpRequest, .pResponse)
    Quit ..constructResponse(.tHttpResponse,.pResponse)
}

Method DownloadFile(pRequest As Ens.StringContainer, Output pDownloadResponse As Ens.StreamContainer) As %Status
{
    
    set binaryStream =##Class(%Stream.FileBinary).%New()
    s tSC=$$$OK
    Set tHttpRequest=##class(%Net.HttpRequest).%New()
    set tHttpRequest.WriteRawMode=1
    Set tHttpRequest.Https  = 1
    do tHttpRequest.SetHeader("Accept","application/json")
    do tHttpRequest.SetHeader("Content-Type","application/json")
    //do tHttpRequest.SetHeader("Connection","keep-alive")
    Set tSC=..AddAccessToken(.tHttpRequest)
    Set ..Adapter.URL="/sites/"_$$$URLENCODE(..SharepointSitename)_"/_api/web/GetFolderByServerRelativeUrl('"_$$$URLENCODE(..SharepointFilePath)_"')/Files('"_$$$URLENCODE(pRequest.StringValue)_"')/OpenBinaryStream()"
    Set tHttpResponse = ##class(%Net.HttpResponse).%New()
	S send="GET"
    $$$TRACE("Getting file from "_..Adapter.URL)
    s tSC = ..SendRequest(.tHttpResponse,send,tHttpRequest, .pResponse)
    //Set tSC=..Adapter.SendFormDataArray(.tHttpResponse, send, tHttpRequest)
    Do tHttpResponse.Data.Rewind()

    while 'tHttpResponse.Data.AtEnd {
		       	set tSC=binaryStream.Write(tHttpResponse.Data.Read(,.tSC1))
		       	if 'tSC1 quit
		    }
    set pDownloadResponse =##Class(Ens.StreamContainer).%New(binaryStream)
    set pDownloadResponse.OriginalFilename=pRequest.StringValue

    Q tSC
}

Method AddFile(pFileToUpload As Ens.StreamContainer, Output pResponse As EnsLib.HTTP.GenericMessage) As %Status
{
    $$$TRACE("AddFile")
	s fn=$$$URLENCODE(##Class(%File).GetFilename(pFileToUpload.OriginalFilename))
    Set tHttpRequest=##class(%Net.HttpRequest).%New()
    set tHttpRequest.WriteRawMode=1
    do tHttpRequest.SetHeader("Accept","application/json;odata=verbose")
    do tHttpRequest.SetHeader("Content-Type","application/json")
    Set ..Adapter.URL="/sites/"_$$$URLENCODE(..SharepointSitename)_"/_api/web/GetFolderByServerRelativeUrl('"_$$$URLENCODE(..SharepointFilePath)_"')/Files/add(url='"_fn_"',overwrite="_$$$URLENCODE(..OverwriteExistingFile)_")"
    Set tSC=..AddAccessToken(.tHttpRequest)
   s tHttpRequest.EntityBody=##Class(%Stream.FileBinary).%New()
	s sc=tHttpRequest.EntityBody.CopyFromAndSave(pFileToUpload.Stream)
	Set tHttpResponse = ##class(%Net.HttpResponse).%New()
	S send="POST"
    $$$TRACE("Sending to "_..Adapter.URL)
	s tSC = ..SendRequest(.tHttpResponse,send,tHttpRequest, .pResponse)
    Quit ..constructResponse(.tHttpResponse,.pResponse)
}

Method SendRequest(Output tHttpResponse As %Net.HttpResponse, send As %String, tHttpRequest As %Net.HttpRequest, Output pResponse As EnsLib.HTTP.GenericMessage) As %Status
{
    s tSC=$$$OK
    
    Set tSC=..Adapter.SendFormDataArray(.tHttpResponse, send, tHttpRequest)
    $$$TRACE("AFTER THE SEND ARRAY")
     $$$TRACE(tHttpResponse)
     //If there was an error and there was a http response with size then display the error output it along with the data
     If $$$ISERR(tSC)&&$IsObject(tHttpResponse)&&$IsObject(tHttpResponse.Data)&&tHttpResponse.Data.Size 
        {
            $$$TRACE(tHttpResponse.Data.Read())
            Set tSC = $$$ERROR($$$EnsErrGeneral,$$$StatusDisplayString(tSC)_":"_tHttpResponse.Data.Read())
        }
        Quit:$$$ISERR(tSC)

    $$$TRACE(tHttpResponse.StatusCode)
    //i think there was a quicker way to do the 200 example i'll have a look MOR 20241204
	if (tHttpResponse.StatusCode<300)
    {
		s tSC=$$$OK
	}
	else //not a 200 response so report back the error 
	{
		set message=tHttpResponse.StatusCode_": "_tHttpResponse.StatusLine_". "
        if $IsObject(tHttpResponse.Data) 
        {
            Try {
                while 'tHttpResponse.Data.AtEnd {
		       	set message=message_tHttpResponse.Data.Read(,.tSC1) 

		    }
            }
            Catch ex {
                //Not bothered as then 
            }
	       	$$$TRACE(message)
		}
		
	}
        $$$TRACE("Constructing Response")
  
    Q tSC
}

Method constructResponse(pHttpResponse As %Net.HttpResponse, Output pResponse As EnsLib.HTTP.GenericMessage) As %Status [ Internal ]
{
    $$$TRACE("construct response")
	Set tSC=$$$OK
    Set tStream=##class(%GlobalBinaryStream).%New()
    if '..CondensedHTMLResponse{
	If $IsObject(pHttpResponse.Data) {
		Set tStream=pHttpResponse.Data
	} Else {
		
		Set tSC=tStream.Write(pHttpResponse.Data)  Quit:$$$ISERR(tSC) tSC
	}
	Set pResponse=##class(EnsLib.HTTP.GenericMessage).%New(tStream,,pHttpResponse)
    }
    else{
    Set pResponse=##class(EnsLib.HTTP.GenericMessage).%New()
        }
  
    set tSC= pResponse.Stream.CopyFromAndSave(tStream)
    
    Do pResponse.HTTPHeaders.SetAt(pHttpResponse.StatusCode,"StatusCode")
	Do pResponse.HTTPHeaders.SetAt(pHttpResponse.StatusLine,"StatusLine")
	Quit tSC
}

/// Assemble Envelope and Body into HTTP EntityBody stream
Method AddAccessToken(ByRef pHttpRequest As %Net.HttpRequest) As %Status
{
	$$$TRACE("SRFT.HttpOperation:AddAccessToken") 
    Set tSC = $$$OK
    #dim error as %OAuth2.Error
    Try {
        
        ; sessionId has a 50 charactor limit
        ; This should support multiple requests in the same production session  
        //Set sessionId=..%SessionId 
        Set sessionId=..OAuthTokenSessionId
        Set isAuthorised = ##class(%SYS.OAuth2.AccessToken).IsAuthorized(..OAuthClientApplicationName,sessionId,..OAuthScope,.accessToken,,.responseProperties,.error)
        ;Throw:'$$$NULLOREF=error ..ExceptionFromOAuth2Error(error)
        Throw:$isobject(error) ..ExceptionFromOAuth2Error(error)
        If 'isAuthorised {
	        $$$TRACE("SRFT.DCCHttpOperation:AddAccessToken isAuthorised false")
            //Add any Parameters
            s paramsarr = [].%FromJSON(..Params)
            s iterator = paramsarr.%GetIterator()
            s properties=""
            While iterator.%GetNext(.key,.value)
            {
                s properties(key)=value
            }

            Set tSC = ##class(%SYS.OAuth2.Authorization).GetAccessTokenClient(..OAuthClientApplicationName,..OAuthScope,.properties,.error,.sessionId)    
            $$$TRACE("SRFT.DCCHttpOperation:AddAccessToken isAuthorised " _ tSC)
            Throw:$isobject(error) ..ExceptionFromOAuth2Error(error)
        }
        ; sendType is the mode of sending access token to resource server
        ; The default is "header" 
        Set sendType = "header"
        ;The default for sslConfiguration comes from the OAuth2.Client instance.        
        Set tSC  = ##class(%SYS.OAuth2.AccessToken).AddAccessToken(pHttpRequest,sendType,,..OAuthClientApplicationName,sessionId)

    } Catch exception {
        Set tSC = exception.AsStatus()
        $$$TRACE("EXCEPTION")
        $$$TRACE("SRFT.DCCHttpOperation:AddAccessToken Code" _ exception.Code)
        $$$TRACE("SRFT.DCCHttpOperation:AddAccessToken Location" _ exception.Location)
        $$$TRACE("SRFT.DCCHttpOperation:AddAccessToken Data" _ exception.Data)
    }

    Return tSC
}

Method ExceptionFromOAuth2Error(pError As %OAuth2.Error) As %Exception.StatusException
{
    Set errorText = pError.AsString()
    $$$LOGERROR(errorText)
    Set tSC = $$$ERROR($$$GeneralError,errorText)
    Return ##class(%Exception.StatusException).CreateFromStatus(tSC)
    ;Or 
    ;If you wish to create one with %New then the 4th argument is a $lb of data values to the error %Status, e.g.
    ;Set exception = ##class(%Exception.StatusException).%New(Name,Code,Location,$lb(arg1,arg2,arg3,arg4))
}

XData MessageMap
{
<MapItems>
  <MapItem MessageType="Ens.StreamContainer">
    <Method>AddFile</Method>
  </MapItem>
  <MapItem MessageType="Messages.Generic.GenericDocument">
    <Method>DeleteFile</Method>
  </MapItem>
  <MapItem MessageType="Ens.StringContainer">
    <Method>DownloadFile</Method>
  </MapItem>
  <MapItem MessageType="Ens.Request">
    <Method>GetFileList</Method>
  </MapItem>
</MapItems>
}

}
