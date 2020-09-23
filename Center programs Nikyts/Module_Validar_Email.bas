Attribute VB_Name = "Module_Validar_Email"
Option Explicit

Private Const Dominios As String = "AERO BIZ COM COOP EDU GOV INFO INT MIL MUSEUM NAME NET ORG PRO " & _
                                    "AC AD AE AF AG AI AL AM AN AO AQ AR AS AT AU AW AZ BA BB BD " & _
                                    "BE BF BG BH BI BJ BM BN BO BR BS BT BV BW BY BZ CA CC CD CF " & _
                                    "CG CH CI CK CL CM CN CO CR CU CV CX CY CZ DE DJ DK DM DO DZ " & _
                                    "EC EE EG EH ER ES ET FI FJ FK FM FO FR GA GD GE GF GG GH GI " & _
                                    "GL GM GN GP GQ GR GS GT GU GW GY HK HM HN HR HT HU ID IE IL " & _
                                    "IM IN IO IQ IR IS IT JE JM JO JP KE KG KH KI KM KN KP KR KW " & _
                                    "KY KZ LA LB LC LI LK LR LS LT LU LV LY MA MC MD MG MH MK ML " & _
                                    "MM MN MO MP MQ MR MS MT MU MV MW MX MY MZ NA NC NE NF NG NI " & _
                                    "NL NO NP NR NU NZ OM PA PE PF PG PH PK PL PM PN PR PS PT PW " & _
                                    "PY QA RE RO RU RW SA SB SC SD SE SG SH SI SJ SK SL SM SN SO " & _
                                    "SR ST SV SY SZ TC TD TF TG TH TJ TK TM TN TO TP TR TT TV TW " & _
                                    "TZ UA UG UK UM US UY UZ VA VC VE VG VI VN VU WF WS YE YT YU " & _
                                    "ZA ZM ZW"

Public Function IsEmail(ByVal Email As String) As Boolean

Dim w        As Integer
Dim sLetra   As String
Dim sSplit() As String
     
    On Error Resume Next
    
    If Len(Email) > 0 Then
        
        If UBound(Split(Email, "@")) <> 1 Or InStr(Email, ".") = 0 Then
            Exit Function
        End If
        
        If Left$(Email, 1) = "@" Or Mid$(Email, Len(Email), 1) = "@" Or InStr(Email, "@.") Or InStr(Email, ".@") Then
            Exit Function
        End If

        If Left$(Email, 1) = "." Or Mid$(Email, Len(Email), 1) = "." Or InStr(Email, "..") Then
            Exit Function
        End If
        
        For w = 1 To Len(Email)
            sLetra = Mid$(Email, w, 1)
            If Not (LCase$(sLetra) Like "[a-z]" Or sLetra = "@" Or sLetra = "." Or sLetra = "-" Or sLetra = "_" Or IsNumeric(sLetra)) Then
                Exit Function
            End If
        Next w
        
        sSplit = Split(UCase$(Trim$(Email)), ".")

        If InStr(Dominios, sSplit(UBound(sSplit))) = 0 Then
            Exit Function
        End If
        
        IsEmail = True
    End If
   
   On Error GoTo 0
   
End Function



