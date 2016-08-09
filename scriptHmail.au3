#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Res_Fileversion=1.0.0
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <HmailUDF.au3>

; constantes de configuration
;================================================
Local $adminName = "Administrator"
Local $adminPassword = "Donjon2016"
Global $bindDomain = "admc.com"
Global $domainShortName = "ADMC"
;================================================

; comportement
;--------------------------------------------------------------------------
; authentification sur le serveur Hmail
Local $objDomainName = _Auth_Hmail($adminName, $adminPassword, $bindDomain)
; procédure de récupération et d'ajout des utilisateurs depuis l'AD
_Add_Users($objDomainName, $domainShortName, $bindDomain)
;--------------------------------------------------------------------------
