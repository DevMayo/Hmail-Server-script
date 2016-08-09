; #INDEX# =======================================================================================================================
; Title .........: HmailUDF
; AutoIt Version : 3.3.14.2
; Description ...: Fonctions permettant de configurer automatiquement un serveur Hmail
; Author(s) .....: DevMayo
; Based on ......: rwgt0su's vbscript Hmail.vbs
; Works on .....: hMailServer 5.6.5


; #CONTENT# =====================================================================================================================
; _Auth_Hmail
; _Add_Users
; _Find_Users
; _Save_Account
; _Account_Exists
;=============================================================================================================

; #FUNCTION# ====================================================================================================================
; Name...........: _Auth_Hmail
; Description ...:  authentifie le script sur Hmail et retourne un objet Hmail chargé avec le nom de domaine
; Syntax.........: _Auth_Hmail($p_adminName, $p_adminPassword, $p_bindDomain)
; Parameters ....:
;		$p_adminName - login d'administration autorisé sur le serveur Hmail
;		$p_adminPassword - mot de passe d'administration
;		$p_bindDomain - nom de domaine
; Return values .: objet COM correspondant au domaine désigné du hMailServer
; ===============================================================================================================================

Func _Auth_Hmail($p_adminName, $p_adminPassword, $p_bindDomain)
	Local $objHMail = ObjCreate("hMailServer.Application")
	$objHMail.Authenticate($p_adminName, $p_adminPassword)
	Local $objDomainName = $objHMail.Domains.ItemByName($p_bindDomain)
	Return $objDomainName
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _Add_Users
; Description ...:  crée une liaison à l'Active Directory et appelle la procédure de recherche des utilisateurs
; Syntax.........: _Add_Users($p_objDomainName, $p_shortName, $p_bindDomain)
; Parameters ....:
;		$p_objDomainName - objet COM correspondant à un domaine du hMailServer
;		$p_shortName - nom court du domaine
;		$p_bindDomain - nom de domaine
; Return values .: None
; ===============================================================================================================================
Func _Add_Users($p_objDomainName, $p_shortName, $p_bindDomain)
	Local $objRoot = ObjGet("LDAP://RootDSE")
	Local $DNC = $objRoot.Get("DefaultNamingContext")
	Local $objDomain = ObjGet("LDAP://" & $DNC)
	_Find_Users($objDomain, $p_objDomainName, $p_shortName, $p_bindDomain)
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _Find_Users
; Description ...:  ; trouve les utilisateurs dans l'active directory qui n'ont pas encore de compte sur Hmail
; Syntax.........: _Find_Users($p_objDomain, $p_objDomainName, $p_shortName, $p_bindDomain)
; Parameters ....:
;		$p_objDomain - object COM correspondant à l'Active Directory
;		$p_objDomainName - objet COM correspondant à un domaine du hMailServer
;		$p_shortName - nom court du domaine
;		$p_bindDomain - nom de domaine
; Return values .: None
; ===============================================================================================================================

Func _Find_Users($p_objDomain, $p_objDomainName, $p_shortName, $p_bindDomain)
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

; #FUNCTION# ====================================================================================================================
; Name...........: _Save_Account
; Description ...:  ; créé un compte sur Hmail en fonction du compte Active Directory
; Syntax.........: _Save_Account($p_ADName, $p_objDomainName, $p_shortName, $p_bindDomain)
; Parameters ....:
;		$p_ADName - string contenant le sam-account-name d'un utilisateur
;		$p_objDomainName - objet COM correspondant à un domaine du hMailServer
;		$p_shortName - nom court du domaine
;		$p_bindDomain - nom de domaine
; Return values .: None
; ===============================================================================================================================

Func _Save_Account($p_ADName, $p_objDomainName, $p_shortName, $p_bindDomain)
	Local $objAccount = $p_objDomainName.Accounts.Add
	$objAccount.Address = $p_ADName & "@" & $p_bindDomain
	$objAccount.Active = True
	$objAccount.MaxSize = 100
	$objAccount.IsAd = True
	$objAccount.ADUsername = $p_ADName
	$objAccount.ADDomain = $p_shortName
	$objAccount.Save
EndFunc

; #FUNCTION# ====================================================================================================================
; Name...........: _Account_Exists
; Description ...:  ; détermine si un compte est déjà présent sur le hMailServer (bool)
; Syntax.........: _Account_Exists($p_ADAccountName, $p_objDomainName)
; Parameters ....:
;		$p_ADAccountName - string contenant le sam-account-name d'un utilisateur
;		$p_objDomainName - objet COM correspondant à un domaine du hMailServer
; Return values .: True si le compte existe déjà, False sinon
; ===============================================================================================================================

Func _Account_Exists($p_ADAccountName, $p_objDomainName)
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