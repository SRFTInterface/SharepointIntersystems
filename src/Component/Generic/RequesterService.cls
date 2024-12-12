/// Standard Service that sends a blank Ens Request out on schedule or interval
Class Component.Generic.RequesterService Extends Ens.BusinessService
{

Parameter ADAPTER = "Ens.InboundAdapter";

/// Configuration item(s) to which to send messages
Property TargetConfigNames As %String(MAXLEN = 2000);

Parameter SETTINGS = "TargetConfigNames:Basic:selector?multiSelect=1&context={Ens.ContextSearch/ProductionItems?targets=1&productionName=@productionId},SearchTableClass::selector?context={Ens.ContextSearch/SearchTableClasses?host=EnsLib.HL7.Service.Standard},LocalFacilityApplication,Framing:Connection:selector?context={Ens.ContextSearch/getDisplayList?host=@currHostId&prop=Framing},AckMode,UseAckCommitCodes,IgnoreInboundAck,AddNackERR,NackErrorCode,BatchHandling,MessageSchemaCategory:Basic:selector?context={Ens.ContextSearch/SchemaCategories?host=EnsLib.HL7.Service.Standard},DefCharEncoding::selector?context={Ens.ContextSearch/CharacterSets}";

ClassMethod OnGetConnections(Output pArray As %String, pItem As Ens.Config.Item)
{
	Do ##super(.pArray,pItem)
	If pItem.GetModifiedSetting("TargetConfigNames",.tValue) {
		For i=1:1:$L(tValue,",") { Set tOne=$ZStrip($P(tValue,",",i),"<>W")  Continue:""=tOne  Set pArray(tOne)="" }
	}
}

Method OnProcessInput(pInput As %RegisteredObject, Output pOutput As %RegisteredObject) As %Status
{
	set req = ##class(Ens.Request).%New()
	Set sc = ..SendRequestAsync(..TargetConfigNames, req, .pResponse)

	Quit sc
}

}
