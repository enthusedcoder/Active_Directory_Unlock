#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#AutoIt3Wrapper_Change2CUI=y
#AutoIt3Wrapper_Res_SaveSource=y
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Res_requestedExecutionLevel=requireAdministrator
#AutoIt3Wrapper_Add_Constants=n
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.15.0 (Beta)
 Author:         myName

 Script Function:
	Template AutoIt script.

#ce ----------------------------------------------------------------------------

; Script Start - Add your code below here
#include <IE.au3>


$store = StringTrimLeft ( $CmdLine[1], 2 )
If FileExists ( @AppDataDir & "\" & $CmdLine[1] & ".txt" ) Then
	Do
		FileDelete ( @AppDataDir & "\" & $CmdLine[1] & ".txt" )
	Until Not FileExists ( @AppDataDir & "\" & $CmdLine[1] & ".txt" )
EndIf

$oIE = _IECreate ( "http://orion/Orion/NetPerfMon/Resources/NodeSearchResults.aspx?Property=Caption&SearchText=" & $store, 0, 0 )
_IELoadWait ( $oIE )
$oForm = _IEGetObjById ( $oIE, "aspnetForm" )
$oLogin = _IEFormElementGetObjByName ( $oForm, "ctl00$BodyContent$Username" )
$oPassword = _IEFormElementGetObjByName ( $oForm, "ctl00$BodyContent$Password" )
_IEFormElementSetValue ( $oLogin, "guest" )
_IEFormElementSetValue ( $oPassword, "@pollo13" )
_IEFormSubmit ( $oForm )
_IELoadWait ( $oIE )
$text = _IEBodyReadText ( $oIE )
_IEQuit ( $oIE )
If StringInStr ( $text, "Megapath" ) > 0 Or StringInStr ( $text, "GTT" ) > 0 Then
	FileWrite ( @AppDataDir & "\" & $CmdLine[1] & ".txt", "Megapath" )
ElseIf StringInStr ( $text, "GRANITE" ) > 0 Then
	FileWrite ( @AppDataDir & "\" & $CmdLine[1] & ".txt", "GRANITE" )
Else
	ConsoleWrite ( "Either Orion does not list the service provider or it is one I have not seen.  If provider is listed, I have placed the information I require in your clipboard.  Please paste it into an email addressed to me and I will add it." )
	ClipPut ( $text )
	Sleep ( 6000 )
EndIf

