/// A class for dealing with filelists from SPO - SharePointOnline.
/// The class firsts gets a list of files and downloads before uploading to a nother folder location 
Class Component.SharePointOnline.Processes.SPOProcessManager Extends Ens.BusinessProcess
{

Parameter SETTINGS = "SharePointService:SharepointService,UploadFileToService:SharepointService,DeleteFileAfterDownload:SharepointService";

/// The service to send all API requests to
Property SharePointService As %String [ InitialExpression = "SharepointRESTConnector" ];

/// the file adapter to upload all files too, after downloaded from SPO
Property UploadFileToService As %String [ InitialExpression = "SharepointDownloadFileEndpoint" ];

/// Default is don't delete file from Sharepoint site after, if true the file will be deleted from SPO.
Property DeleteFileAfterDownload As %Boolean [ InitialExpression = 0, Required ];

ClassMethod OnGetConnections(Output pArray As %String, pItem As Ens.Config.Item)
{
	d pItem.PopulateModifiedSettings()
	s (tValue,tIndex)=""
	for
	{
		s tSetting = pItem.ModifiedSettings.GetNext(.tIndex) q:tIndex=""
		if tSetting.Name = "SharePointService" s tValue=tSetting.Value  q:tIndex=""
		If tSetting.Name = "UploadFileToService" Set tValue=tValue_","_tSetting.Value q:tIndex=""
	}
	for i=1:1:$Length(tValue,",") 
	{
		s tOne=$ZStrip($p(tValue,",",i),"<>W")  Continue:""=tOne  s pArray(tOne)=""
	}
}

Method OnRequest(pRequest As Ens.Request, Output pResponse As EnsLib.HTTP.GenericMessage) As %Status
{

	s response=##Class(EnsLib.HTTP.GenericMessage).%New()
	set tSC=..SendRequestSync(..SharePointService,pRequest,.response,60)  Quit:$$$ISERR(tSC)
	s d=""
	set result = [].%FromJSON(response.StreamCF)
	s resa= [].%FromJSON(result.d.%ToJSON())
	s resb=[].%FromJSON(resa.results.%ToJSON())
	s iter=resb.%GetIterator()
	s i=0
	while iter.%GetNext(.key,.val) 
	{
		$$$TRACE("Download and upload "_resb.%Get(i).Name)
		//download file
		Try
		{
			s str=##Class(Ens.StringContainer).%New(resb.%Get(i).Name)
			set tSC=..SendRequestSync(..SharePointService,str,.response,60)  Quit:$$$ISERR(tSC)
			//upload file
			if tSC=$$$OK {s tSC=..SendRequestSync(..UploadFileToService, response,60)}
			if ..DeleteFileAfterDownload
			{
				s dao=##Class(Messages.Generic.GenericDocument).%New()
				s dao.FileName = str.StringValue
				s tSC=..SendRequestSync(..SharePointService,dao, response,60)
			}
			s i=i+1
		}
		Catch ex 
		{
			Set tSC = ex.AsStatus()
		}
	}
	Quit tSC
}

Storage Default
{
<Data name="DownloadFileProcessDefaultData">
<Subscript>"DownloadFileProcess"</Subscript>
<Value name="1">
<Value>SharePointService</Value>
</Value>
<Value name="2">
<Value>UploadFileToService</Value>
</Value>
<Value name="3">
<Value>DeleteFileAfterDownload</Value>
</Value>
</Data>
<DefaultData>DownloadFileProcessDefaultData</DefaultData>
<Type>%Storage.Persistent</Type>
}

}
