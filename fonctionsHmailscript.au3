Func _Auth_Hmail($p_adminName, $p_adminPassword, $p_bindDomain) ; authentifie le script sur Hmail et retourne un objet Hmail chargé avec le nom de domaine
	Local $objHMail = ObjCreate("hMailServer.Application")
	$objHMail.Authenticate($p_adminName, $p_adminPassword)
	;chargement du nom de domaine sur lequel on ajoute un compte
	Local $objDomainName = $objHMail.Domains.ItemByName($p_bindDomain)
	Return $objDomainName
EndFunc

Func _Add_Users($p_objDomainName, $p_shortName, $p_bindDomain) ; trouve et ajoute les utilisateurs depuis l'active directory
	; liaison au domaine
	Local $objRoot = ObjGet("LDAP://RootDSE")
	Local $DNC = $objRoot.Get("DefaultNamingContext")
	Local $objDomain = ObjGet("LDAP://" & $DNC)
	_Find_Users($objDomain, $p_objDomainName, $p_shortName, $p_bindDomain)
EndFunc

Func _Find_Users($p_objDomain, $p_objDomainName, $p_shortName, $p_bindDomain) ; trouve les utilisateurs dans l'active directory qui n'ont pas encore de compte sur Hmail
	Local $samAccountName
	If IsObj($p_objDomain) Then
		For $objMember In $p_objDomain
			If $objMember.Class = "user" Then
				$samAccountName = $objMember.samAccountName
				If Not _Account_Exists($samAccountName, $p_objDomainName) Then
					_Save_Account($samAccountName, $p_objDomainName, $p_shortName, $p_bindDomain)
				EndIf
				$samAccountName = "-"
			ElseIf $objMember.Class = "organizationalUnit" Or $objMember.Class = "container" Then
				_Find_Users($objMember, $p_objDomainName, $p_shortName, $p_bindDomain)
			EndIf
		Next
	EndIf
EndFunc

Func _Save_Account($p_ADName, $p_objDomainName, $p_shortName, $p_bindDomain); créé un compte sur Hmail en fonction du compte Active Directory
	Local $objAccount = $p_objDomainName.Accounts.Add
	$objAccount.Address = $p_ADName & "@" & $p_bindDomain
	$objAccount.Active = True
	$objAccount.MaxSize = 100
	$objAccount.IsAd = True
	$objAccount.ADUsername = $p_ADName
	$objAccount.ADDomain = $p_shortName
	$objAccount.Save
EndFunc

Func _Account_Exists($p_ADAccountName, $p_objDomainName); détermine si un compte est déjà présent dans la bdd Hmail (bool)
	Local $exists = False
	Local $objAccount = $p_objDomainName.Accounts
	Local $count = $objAccount.Count
	For $i = 0 To $count -1
		If $objAccount.Item($i).ADUsername = $p_ADAccountName Then
			$exists = True
		EndIf
	Next
	Return $exists
EndFunc