; constantes de configuration
Local $adminName = "Administrator"
Local $adminPassword = "Donjon2016"
Global $bindDomain = "admc.com"
Global $domainShortName = "ADMC"

Global $objAccount
Global $objHMail = ObjCreate("hMailServer.Application")
$objHMail.Authenticate($adminName, $adminPassword)

;chargement du nom de domaine sur lequel on ajoute un compte
Local $objDomainName = $objHMail.Domains.ItemByName($bindDomain)
_Add_Users()

Func _Add_Users()
	; liaison au domaine
	Local $objRoot = ObjGet("LDAP://RootDSE")
	Local $DNC = $objRoot.Get("DefaultNamingContext")
	Local $objDomain = ObjGet("LDAP://" & $DNC)
	_Find_Users($objDomain)
EndFunc

Func _Find_Users($p_objDomain)
;	Local $objwb
;	Local $x
	Local $samAccountName
	If IsObj($p_objDomain) Then
		For $objMember In $p_objDomain
			If $objMember.Class = "user" Then
				;$objwb.Cells($x, 1).Value = $objMember.Class
				$samAccountName = $objMember.samAccountName
				If Not _Account_Exists($samAccountName) Then
					_Save_Account($samAccountName)
				EndIf
				$samAccountName = "-"
			EndIf

			If $objMember.Class = "organizationalUnit" Or $objMember.Class = "container" Then
				_Find_Users($objMember)
			EndIf
		Next
	EndIf
EndFunc

Func _Save_Account($p_ADName)
	$objAccount = $objDomainName.Accounts.Add
	$objAccount.Address = $p_ADName & "@" & $bindDomain
	$objAccount.Active = True
	$objAccount.MaxSize = 100
	$objAccount.IsAd = True
	$objAccount.ADUsername = $p_ADName
	$objAccount.ADDomain = $domainShortName
	$objAccount.Save
EndFunc

Func _Account_Exists($p_ADAccountName); détermine si un compte est déjà présent dans la bdd Hmail (bool)
	Local $exists = False
	$objAccount = $objDomainName.Accounts
	Local $count = $objAccount.Count
	For $i = 0 To $count -1
		If $objAccount.Item($i).ADUsername = $p_ADAccountName Then
			$exists = True
		EndIf
	Next
	Return $exists
EndFunc