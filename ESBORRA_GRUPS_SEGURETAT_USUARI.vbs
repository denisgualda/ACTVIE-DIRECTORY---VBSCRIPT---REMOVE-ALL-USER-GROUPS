'**********************************************************************
'ESBORRA GRUPS D'USUARI
'**********************************************************************

''**********************************************************************
'RECORRE TOTS ESL GRUPS I ESBORRA MENTRE OBTE CADA GRUP
'ES NECESITA OBTENIR CN,OU,DC D'USUARI I GRUP
	'es pot obtenir la informacio amb la comanda:-->  dsquery group/user -name USUARI/GRUP

''**********************************************************************

Dim ObjUser, ObjRootDSE, ObjConn, ObjRS 
Dim GroupCollection, ObjGroup 
Dim StrUserName, StrDomName, StrSQL 

'ESBORRA GRUPS
	
dim groupPath
dim userPath
dim member

Set ObjRootDSE = GetObject("LDAP://RootDSE") 
StrDomName = Trim(ObjRootDSE.Get("DefaultNamingContext")) 
Set ObjRootDSE = Nothing 

' -- ENTRA EL DNI DE L'USUARI
strUserName = InputBox("Nom d'usuari o DNI(en domini GCB)") 
StrSQL = "Select ADsPath From 'LDAP://" & StrDomName & "' Where ObjectCategory = 'User' AND SAMAccountName = '" & StrUserName & "'" 
 
Set ObjConn = CreateObject("ADODB.Connection") 
ObjConn.Provider = "ADsDSOObject":    ObjConn.Open "Active Directory Provider" 
Set ObjRS = CreateObject("ADODB.Recordset") 
ObjRS.Open StrSQL, ObjConn 
If Not ObjRS.EOF Then 
    ObjRS.MoveLast:    ObjRS.MoveFirst 
    WScript.Echo vbNullString 
    Set ObjUser = GetObject (Trim(ObjRS.Fields("ADsPath").Value)) 
    Set GroupCollection = ObjUser.Groups 
    For Each ObjGroup In GroupCollection 
    '------------------------------------------------------
        'AFEGIT PER UTILIZAR LA FUNCIO: GET DISTINGUESHED NAME
        set objSystemInfo = CreateObject("ADSystemInfo") 
        strDomain = objSystemInfo.DomainShortName

			'Converteix a distingueshed name el nom d'usuari.
			strGroupPath = GetUserDN(objGroup.Cn,strDomain)
			groupPath = "LDAP://" & strGroupPath
			'wscript.echo groupPath

		'Converteix a distingueshed name el nom d'usuari.
			struserPath = GetUserDN(strUserName,strDomain)
			userPath = "LDAP://" & struserPath
			'wscript.echo userPath	

		'CRIDA LA FUNCIONA PER ESBORRAR GRUP
		removeFromGroup userPath,groupPath
	'------------------------------------------------------
		
		'Escriu per pantalla grups ESBORRATS (objGroup.CN)
		WScript.Echo "  GRUP ESBORRAT: " &  ObjGroup.CN 
	Next 
	Set ObjGroup = Nothing:    Set GroupCollection = Nothing:    Set ObjUser = Nothing 
Else 
    WScript.Echo "L'usuari: " & StrUserName & " no s'ha trobat al domini" 
End If 
ObjRS.Close:    Set ObjRS = Nothing 
ObjConn.Close:    Set ObjConn = Nothing 


'**************************************************************
'FUNCIO ELIMINA GRUP
sub removeFromGroup(userPath, groupPath)

	dim objGroup
	set objGroup = getobject(groupPath)
	
	for each member in objGroup.members
		if lcase(member.adspath) = lcase(userPath) then
			objGroup.Remove(userPath)
			exit sub
		end if
	next
end sub
'**************************************************************


'**************************************************************
'FUNCIO OBTENIR DISTINGUESHED NAME
Function GetUserDN(byval strUserName,byval strDomain)

	Set objTrans = CreateObject("NameTranslate")
	objTrans.Init 1, strDomain
	objTrans.Set 3, strDomain & "\" & strUserName 
	strUserDN = objTrans.Get(1) 
	GetUserDN = strUserDN

End function
'**************************************************************