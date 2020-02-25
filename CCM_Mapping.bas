Attribute VB_Name = "CCM_Mapping"
Option Explicit
Option Private Module       ' Excel users cannot access procedures

'*********LIST OF Comp2Find Values
'PRODUCE MAPPING
'        Comp2Find = "ColesWNAT1"    -   C_Code
'        Comp2Find = "ColesWNAT2"    -   C_CCode
'        Comp2Find = "ColesWNAT3"    -   C_CCode2
'        Comp2Find = "ColesWNAT4"    -   C_CCode3
'        Comp2Find = "ColesWNSW"    -   C_Code2
'        Comp2Find = "ColesWVIC"    -   C_Code3
'        Comp2Find = "ColesWQLD"    -   C_Code4
'        Comp2Find = "ColesWSA"     -   C_Code5
'        Comp2Find = "ColesWWA"     -   C_Code6
'        Comp2Find = "WWWeb"        -   W_Code
'        Comp2Find = "WWWNAT1"        -   W_Code
'        Comp2Find = "WWWNAT2"        -   W_HBCode
'        Comp2Find = "WWWNAT3"        -   W_HBCode2
'        Comp2Find = "WWWNAT4"        -   W_HBCode3
'        Comp2Find = "WWWNSW"       -   W_Code2
'        Comp2Find = "WWWVIC"       -   W_Code3
'        Comp2Find = "WWWQLD"       -   W_Code4
'        Comp2Find = "WWWSA"        -   W_Code5
'        Comp2Find = "WWWWA"        -   W_Code6

'CORE RANGE MAPPING COLES
'        Comp2Find = "ColesWeb"     -   C_Code
'        Comp2Find = "ColesSB1"      -   C_SBCode
'        Comp2Find = "ColesSB2"     -   C_SBCode2
'        Comp2Find = "ColesSB3"     -   C_SBCode3
'        Comp2Find = "ColesPL1"   -   C_CCode
'        Comp2Find = "ColesPL2"  -   C_CCode2
'        Comp2Find = "ColesPL3"  -   C_CCode3
'        Comp2Find = "ColesPL4"  -   C_CCode4
'        Comp2Find = "ColesVal1"     -   C_CVCode
'        Comp2Find = "ColesVal2"    -   C_CVCode2
'        Comp2Find = "ColesVal3"    -   C_CVCode3
'        Comp2Find = "ColesCB1"      -   B_CBCode
'        Comp2Find = "ColesPB1"      -   B_PBCode
'        Comp2Find = "ColesPB2"      -   S_Code3
'        Comp2Find = "ColesML1"     -   B_MLCode
'        Comp2Find = "ColesML2"     -   B_MLCode2
'        Comp2Find = "ColesML3"     -   B_MLCode5



'CORE RANGE MAPPING WW
'        Comp2Find = "WWWeb"        -   W_Code
'        Comp2Find = "WWWW1"        -   W_SelCode
'        Comp2Find = "WWWW2"        -   W_SelCode2
'        Comp2Find = "WWWW3"        -   W_SelCode3
'        Comp2Find = "WWWW4"        -   W_SelCode4
'        Comp2Find = "WWWW5"        -   W_SelCode5
'        Comp2Find = "WWHB1"         -   W_HBCode
'        Comp2Find = "WWHB2"        -   W_HBCode2
'        Comp2Find = "WWHB3"        -   W_HBCode3
'        Comp2Find = "WWCB"         -   B_CBCode3
'        Comp2Find = "WWPB1"         -   B_PBCode3
'        Comp2Find = "WWPB2"         -  S_Code4
'        Comp2Find = "WWML1"        -   B_MLCode3
'        Comp2Find = "WWML2"        -   B_MLCode4
'        Comp2Find = "WWML3"        -   B_MLCode6


'DAN MURPHYS MAPPING
'        Comp2Find = "DM1"          -   DM_Code1 / DM_Code1Pack
'        Comp2Find = "DM2"          -   DM_Code2 / DM_Code2Pack
'        Comp2Find = "DMQ"          -   B_CBCode2 / DM_CBCode2Pack


'FIRST CHOICE MAPPING
'        Comp2Find = "FC1"          -   S_Code11 / S_Code11Pack
'        Comp2Find = "FC2"          -   S_Code12 / S_Code12Pack
'        Comp2Find = "FCQ"          -   B_PBCode2 / B_PBCode2Pack



'OLD MAPPINGS
'        Comp2Find = "DMCB"         -   B_CBCode2 'OLD MAPPING ALCOHOL
'        Comp2Find = "DMPB"         -   B_PBCode2 'OLD MAPPING ALCOHOL
'        Comp2Find = "DMML1"        -   B_MLCode5 'OLD MAPPING ALCOHOL       - Only avaliable for Alcohol Products
'        Comp2Find = "DMML2"        -   B_MLCode6 'OLD MAPPING ALCOHOL       - Only avaliable for Alcohol Products
'        Comp2Find = "BrandAvg"     -   B_AVCode    'OLD MAPPING

'NEW Price Optimization Dicount Corridor Allocation
'        Comp2Find = SCode_7    CompProduct = S_Code8     Competitor = S_Code9








Type MatchTypeData
    ApplyMatchComment As String
    DeactivateMatchComment As String
    DbFieldName As String
    AlcoholPackDBField As String
    Competitor As String
    CompetitorLng As String
    Description As String
    CoreAlcProd As String
    OptionButtonName As String
    MappingTableNumber As Long
    MappingTableNumberPack As Long
    Comp2Find As String
End Type
Function MatchType(ByVal Comp2Find As String, Optional sCompetitor As String, Optional bIsProduce As Boolean) As MatchTypeData
Dim C2F As MatchTypeData

    Select Case Comp2Find
'NEW Price Optimization Dicount Corridor Allocation
'        Comp2Find = "Bench" MatchType/Corridor = SCode_7    CompProduct = S_Code8     Competitor = S_Code9
    Case "Bench"
            If sCompetitor <> "" Then C2F.Competitor = sCompetitor
            If sCompetitor = "C" Then
                C2F.CompetitorLng = "Coles"
                If bIsProduce Then C2F.CoreAlcProd = "Produce" Else C2F.CoreAlcProd = "Core"
            ElseIf sCompetitor = "WW" Then
                C2F.CompetitorLng = "Woolworths"
                If bIsProduce Then C2F.CoreAlcProd = "Produce" Else C2F.CoreAlcProd = "Core"
            ElseIf sCompetitor = "DM" Then
                C2F.CompetitorLng = "Dan Murphys"
                C2F.CoreAlcProd = "Alcohol"
            ElseIf sCompetitor = "FC" Then
                C2F.CompetitorLng = "First Choice"
                C2F.CoreAlcProd = "Alcohol"
            ElseIf sCompetitor = "AZ" Then
                C2F.CompetitorLng = "Amazon"
                C2F.CoreAlcProd = "Core"
            End If
            C2F.Description = Comp2Find
            C2F.DbFieldName = "S_Code8"
            C2F.OptionButtonName = "ob_Bench"
'            C2F.MappingTableNumber = 14
'NEED TO UNDERSTAND AND HANDLE DANS AND FIRST C2F.AlcoholPackDBField = ""
    
    
'CORE RANGE COLES CODES
        Case "ColesWeb"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Watch"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_Code"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_Web"
            C2F.MappingTableNumber = 14
            
            
        Case "ColesSB1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Smartbuy 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_SBCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_SB1"
            C2F.MappingTableNumber = 4
           
            
        Case "ColesSB2"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Smartbuy 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_SBCode2"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_SB2"
            C2F.MappingTableNumber = 5
            
            
        Case "ColesSB3"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Smartbuy 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_SBCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_SB3"
            C2F.MappingTableNumber = 6
            

        Case "ColesVal1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Value 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CVCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_CV1"
            C2F.MappingTableNumber = 11
            
        Case "ColesVal2"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Value 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CVCode2"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_CV2"
            C2F.MappingTableNumber = 12
        
        Case "ColesVal3"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Value 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CVCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_CV3"
            C2F.MappingTableNumber = 13
            
        Case "ColesCB1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Control Brand"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_CBCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_CB"
            C2F.MappingTableNumber = 39
            
        Case "ColesPL1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Private Label 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_PL1"
            C2F.MappingTableNumber = 7
            
        Case "ColesPL2"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Private Label 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CCode2"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_PL2"
            C2F.MappingTableNumber = 8
            
        Case "ColesPL3"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Private Label 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_PL3"
            C2F.MappingTableNumber = 9
            
        Case "ColesPL4"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Private Label 4"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "C_CCode4"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_PL4"
            C2F.MappingTableNumber = 10
        Case "ColesPB1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Phantom Brand 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_PBCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_PB1"
            C2F.MappingTableNumber = 40
        Case "ColesPB2"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Phantom Brand 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "S_Code3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_PB2"
            C2F.MappingTableNumber = 55
        Case "ColesML1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Market Leader 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_MLCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_ML1"
            C2F.MappingTableNumber = 44
        Case "ColesML2"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Market Leader 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_MLCode2"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_ML2"
            C2F.MappingTableNumber = 45
        Case "ColesML3"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "Market Leader 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_MLCode5"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_C_ML3"
            C2F.MappingTableNumber = 48
            
'CORE RANGE WW CODES

        Case "WWWeb"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Watch"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_Code"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_Web"
            C2F.MappingTableNumber = 20
            
        Case "WWHB1"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Homebrand 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_HBCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_HB1"
            C2F.MappingTableNumber = 31
        
        Case "WWHB2"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Homebrand 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_HBCode2"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_HB2"
            C2F.MappingTableNumber = 32
            
        Case "WWHB3"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Homebrand 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_HBCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_HB3"
            C2F.MappingTableNumber = 33
            
        Case "WWWW1"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Private Label 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_SelCode"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_PL1"
            C2F.MappingTableNumber = 26
            
        Case "WWWW2"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Private Label 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_SelCode2"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_PL2"
            C2F.MappingTableNumber = 27
        
        Case "WWWW3"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Private Label 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_SelCode3"
            C2F.CoreAlcProd = "Core"
             C2F.OptionButtonName = "ob_W_PL3"
            C2F.MappingTableNumber = 28
        Case "WWWW4"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Private Label 4"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_SelCode4"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_PL4"
            C2F.MappingTableNumber = 29
        Case "WWWW5"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Private Label 5"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "W_SelCode5"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_PL5"
            C2F.MappingTableNumber = 30
           
        Case "WWCB1"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Control Brand"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_CBCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_CB"
            C2F.MappingTableNumber = 54
            
        Case "WWPB1"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Phantom Brand 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_PBCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_PB1"
            C2F.MappingTableNumber = 43
        Case "WWPB2"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Phantom Brand 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "S_Code4"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_PB2"
            C2F.MappingTableNumber = 56
            
        Case "WWML1"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Market Leader 1"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_MLCode3"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_ML1"
            C2F.MappingTableNumber = 46
        Case "WWML2"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Market Leader 2"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_MLCode4"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_ML2"
            C2F.MappingTableNumber = 47
        Case "WWML3"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "Market Leader 3"
            C2F.AlcoholPackDBField = ""
            C2F.DbFieldName = "B_MLCode6"
            C2F.CoreAlcProd = "Core"
            C2F.OptionButtonName = "ob_W_ML3"
            C2F.MappingTableNumber = 49
 'ALCOHOL

        Case "DM1"
            C2F.Competitor = "DM"
            C2F.CompetitorLng = "Dan Murphys"
            C2F.Description = "Price 1"
            C2F.AlcoholPackDBField = "DM_Code1Pack"
            C2F.DbFieldName = "DM_Code1"
            C2F.CoreAlcProd = "Alcohol"
            C2F.OptionButtonName = "ob_DM1"
            C2F.MappingTableNumber = 34
            C2F.MappingTableNumberPack = 35
        Case "DM2"
            C2F.Competitor = "DM"
            C2F.CompetitorLng = "Dan Murphys"
            C2F.Description = "Price 2"
            C2F.AlcoholPackDBField = "DM_Code2Pack"
            C2F.DbFieldName = "DM_Code2"
            C2F.CoreAlcProd = "Alcohol"
            C2F.OptionButtonName = "ob_DM2"
            C2F.MappingTableNumber = 36
            C2F.MappingTableNumberPack = 37
        
        Case "DMQ"
            C2F.Competitor = "DM"
            C2F.CompetitorLng = "Dan Murphys"
            C2F.Description = "Quality"
            C2F.AlcoholPackDBField = "B_CBCode2Pack"
            C2F.DbFieldName = "B_CBCode2"
            C2F.CoreAlcProd = "Alcohol"
            C2F.OptionButtonName = "ob_DMQ"
            C2F.MappingTableNumber = 52
            C2F.MappingTableNumberPack = 53
            
        Case "FC1"
            C2F.Competitor = "FC"
            C2F.CompetitorLng = "First Choice"
            C2F.Description = "Price 1"
            C2F.AlcoholPackDBField = "S_Code11Pack"
            C2F.DbFieldName = "S_Code11"
            C2F.CoreAlcProd = "Alcohol"
            C2F.OptionButtonName = "ob_FC1"
            C2F.MappingTableNumber = 63
            C2F.MappingTableNumberPack = 64

           
        Case "FC2"
            C2F.Competitor = "FC"
            C2F.CompetitorLng = "First Choice"
            C2F.Description = "Price 2"
            C2F.AlcoholPackDBField = "S_Code12Pack"
            C2F.DbFieldName = "S_Code12"
            C2F.CoreAlcProd = "Alcohol"
            C2F.OptionButtonName = "ob_FC2"
            C2F.MappingTableNumber = 65
            C2F.MappingTableNumberPack = 66
        Case "FCQ"
            C2F.Competitor = "FC"
            C2F.CompetitorLng = "First Choice"
            C2F.Description = "Quality"
            C2F.AlcoholPackDBField = "B_PBCode2Pack"
            C2F.DbFieldName = "B_PBCode2"
            C2F.CoreAlcProd = "Alcohol"
            C2F.OptionButtonName = "ob_FCQ"
            C2F.MappingTableNumber = 41
            C2F.MappingTableNumberPack = 42
            
'Produce
        
        Case "ColesWNAT1"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "National Produce 1"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_Code"
            C2F.OptionButtonName = "ob_C_NAT_Produce1"
            C2F.MappingTableNumber = 14
            
        Case "ColesWNAT2"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "National Produce 2"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_CCode"
            C2F.OptionButtonName = "ob_C_NAT_Produce2"
            C2F.MappingTableNumber = 7
        Case "ColesWNAT3"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "National Produce 3"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_CCode2"
            C2F.OptionButtonName = "ob_C_NAT_Produce3"
            C2F.MappingTableNumber = 8
        Case "ColesWNAT4"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "National Produce 4"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_CCode3"
            C2F.OptionButtonName = "ob_C_NAT_Produce4"
            C2F.MappingTableNumber = 9
        Case "ColesWNSW"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "NSW Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_Code2"
            C2F.OptionButtonName = "ob_C_NSW_Produce"
            C2F.MappingTableNumber = 15
        Case "ColesWVIC"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "VIC Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_Code3"
            C2F.OptionButtonName = "ob_C_VIC_Produce"
            C2F.MappingTableNumber = 16
        Case "ColesWQLD"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "QLD Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_Code4"
            C2F.OptionButtonName = "ob_C_QLD_Produce"
            C2F.MappingTableNumber = 17
        Case "ColesWSA"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "SA Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_Code5"
            C2F.OptionButtonName = "ob_C_SA_Produce"
            C2F.MappingTableNumber = 18
        Case "ColesWWA"
            C2F.Competitor = "C"
            C2F.CompetitorLng = "Coles"
            C2F.Description = "WA Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "C_Code6"
            C2F.OptionButtonName = "ob_C_WA_Produce"
            C2F.MappingTableNumber = 19
        Case "WWWNAT1"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "National Produce 1"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_Code"
            C2F.OptionButtonName = "ob_W_NAT_Produce1"
            C2F.MappingTableNumber = 20
        Case "WWWNAT2"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "National Produce 2"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_HBCode"
            C2F.OptionButtonName = "ob_W_NAT_Produce2"
            C2F.MappingTableNumber = 31
        Case "WWWNAT3"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "National Produce 3"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_HBCode2"
            C2F.OptionButtonName = "ob_W_NAT_Produce3"
            C2F.MappingTableNumber = 32
        Case "WWWNAT4"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "National Produce 4"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_HBCode3"
            C2F.OptionButtonName = "ob_W_NAT_Produce3"
            C2F.MappingTableNumber = 33
        Case "WWWNSW"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "NSW Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_Code2"
            C2F.OptionButtonName = "ob_W_NSW_Produce"
            C2F.MappingTableNumber = 21
        Case "WWWVIC"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "VIC Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_Code3"
            C2F.OptionButtonName = "ob_W_VIC_Produce"
            C2F.MappingTableNumber = 22
        Case "WWWQLD"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "QLD Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_Code4"
            C2F.OptionButtonName = "ob_W_QLD_Produce"
            C2F.MappingTableNumber = 23
        Case "WWWSA"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "SA Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_Code5"
            C2F.OptionButtonName = "ob_W_SA_Produce"
            C2F.MappingTableNumber = 24
        Case "WWWWA"
            C2F.Competitor = "WW"
            C2F.CompetitorLng = "Woolworths"
            C2F.Description = "WA Produce"
            C2F.AlcoholPackDBField = ""
            C2F.CoreAlcProd = "Produce"
            C2F.DbFieldName = "W_Code6"
            C2F.OptionButtonName = "ob_W_WA_Produce"
            C2F.MappingTableNumber = 25

    End Select

    C2F.Comp2Find = Comp2Find
    C2F.ApplyMatchComment = C2F.CompetitorLng & " " & C2F.Description & " Match Linked"
    C2F.DeactivateMatchComment = C2F.CompetitorLng & " " & C2F.Description & " Match Deactivated"


    MatchType = C2F

End Function

Function ob_FindComp(ByVal ob_Name As String) As String

    Select Case ob_Name
    
''Price Optimization Benhcmark
        Case "ob_Bench"
            ob_FindComp = "Bench"



'CORE RANGE COLES CODES
        Case "ob_C_Web"
            ob_FindComp = "ColesWeb"
        Case "ob_C_SB1"
            ob_FindComp = "ColesSB1"
        Case "ob_C_SB2"
            ob_FindComp = "ColesSB2"
        Case "ob_C_SB3"
            ob_FindComp = "ColesSB3"
        Case "ob_C_CV1"
            ob_FindComp = "ColesVal1"
        Case "ob_C_CV2"
            ob_FindComp = "ColesVal1"
        Case "ob_C_CV3"
            ob_FindComp = "ColesVal3"
        Case "ob_C_CB"
            ob_FindComp = "ColesCB1"
        Case "ob_C_PL1"
            ob_FindComp = "ColesPL1"
        Case "ob_C_PL2"
            ob_FindComp = "ColesPL2"
        Case "ob_C_PL3"
            ob_FindComp = "ColesPL3"
        Case "ob_C_PL4"
            ob_FindComp = "ColesPL4"
        Case "ob_C_PB1"
            ob_FindComp = "ColesPB1"
        Case "ob_C_PB2"
            ob_FindComp = "ColesPB2"
        Case "ob_C_ML1"
            ob_FindComp = "ColesML1"
        Case "ob_C_ML2"
            ob_FindComp = "ColesML2"
        Case "ob_C_ML3"
            ob_FindComp = "ColesML3"
            
'CORE RANGE WW CODES

        Case "ob_W_Web"
            ob_FindComp = "WWWeb"
        Case "ob_W_HB1"
            ob_FindComp = "WWHB1"
        Case "ob_W_HB2"
            ob_FindComp = "WWHB2"
        Case "ob_W_HB3"
            ob_FindComp = "WWHB3"
        Case "ob_W_PL1"
            ob_FindComp = "WWWW1"
        Case "ob_W_PL2"
            ob_FindComp = "WWWW2"
        Case "ob_W_PL3"
            ob_FindComp = "WWWW3"
        Case "ob_W_PL4"
            ob_FindComp = "WWWW4"
        Case "ob_W_PL5"
            ob_FindComp = "WWWW5"
        Case "ob_W_CB"
            ob_FindComp = "WWCB1"
        Case "ob_W_ML1"
            ob_FindComp = "WWML1"
        Case "ob_W_ML2"
            ob_FindComp = "WWML2"
        Case "ob_W_ML3"
            ob_FindComp = "WWML3"
        Case "ob_W_PB1"
            ob_FindComp = "WWPB1"
        Case "ob_W_PB2"
            ob_FindComp = "WWPB2"
 'ALCOHOL
        Case "ob_DM1"
            ob_FindComp = "DM1"
        Case "ob_DM2"
            ob_FindComp = "DM2"
        Case "ob_DMQ"
            ob_FindComp = "DMQ"
        Case "ob_FC1"
            ob_FindComp = "FC1"
        Case "ob_FC2"
            ob_FindComp = "FC2"
        Case "ob_FCQ"
            ob_FindComp = "FCQ"
            
'Produce
        Case "ob_C_NAT_Produce1"
            ob_FindComp = "ColesWNAT1"
        Case "ob_C_NAT_Produce2"
            ob_FindComp = "ColesWNAT2"
        Case "ob_C_NAT_Produce3"
            ob_FindComp = "ColesWNAT3"
        Case "ob_C_NAT_Produce4"
            ob_FindComp = "ColesWNAT4"
        Case "ob_C_NSW_Produce"
            ob_FindComp = "ColesWNSW"
        Case "ob_C_VIC_Produce"
            ob_FindComp = "ColesWVIC"
        Case "ob_C_QLD_Produce"
            ob_FindComp = "ColesWQLD"
        Case "ob_C_SA_Produce"
            ob_FindComp = "ColesWSA"
        Case "ob_C_WA_Produce"
            ob_FindComp = "ColesWWA"
        Case "ob_W_NAT_Produce1"
            ob_FindComp = "WWWNAT1"
        Case "ob_W_NAT_Produce2"
            ob_FindComp = "WWWNAT2"
        Case "ob_W_NAT_Produce3"
            ob_FindComp = "WWWNAT3"
        Case "ob_W_NAT_Produce4"
            ob_FindComp = "WWWNAT4"
        Case "ob_W_NSW_Produce"
            ob_FindComp = "WWWNSW"
        Case "ob_W_VIC_Produce"
            ob_FindComp = "WWWVIC"
        Case "ob_W_QLD_Produce"
            ob_FindComp = "WWWQLD"
        Case "ob_W_SA_Produce"
            ob_FindComp = "WWWSA"
        Case "ob_W_WA_Produce"
            ob_FindComp = "WWWWA"
    End Select
'Debug.Print ob_FindComp
End Function
Function CMM_getComp2Find(ByVal DBName As String, ByVal CGno As Long)
If CGno = 58 Or CGno = 27 Then
    Select Case DBName
        Case "C_Code"
            CMM_getComp2Find = "ColesWNAT1"
        Case "C_CCode"
            CMM_getComp2Find = "ColesWNAT2"
        Case "C_CCode2"
            CMM_getComp2Find = "ColesWNAT3"
        Case "C_CCode3"
            CMM_getComp2Find = "ColesWNAT4"
        Case "C_Code2"
            CMM_getComp2Find = "ColesWNSW"
        Case "C_Code3"
            CMM_getComp2Find = "ColesWVIC"
        Case "C_Code4"
            CMM_getComp2Find = "ColesWQLD"
        Case "C_Code5"
            CMM_getComp2Find = "ColesWSA"
        Case "C_Code6"
            CMM_getComp2Find = "ColesWWA"
''        Case "W_Code"
''            CMM_getComp2Find = "WWWeb"
        Case "W_Code"
            CMM_getComp2Find = "WWWNAT1"
        Case "W_HBCode"
            CMM_getComp2Find = "WWWNAT2"
        Case "W_HBCode2"
            CMM_getComp2Find = "WWWNAT3"
        Case "W_HBCode3"
            CMM_getComp2Find = "WWWNAT4"
        Case "W_Code2"
            CMM_getComp2Find = "WWWNSW"
        Case "W_Code3"
            CMM_getComp2Find = "WWWVIC"
        Case "W_Code4"
            CMM_getComp2Find = "WWWQLD"
        Case "W_Code5"
            CMM_getComp2Find = "WWWSA"
        Case "W_Code6"
            CMM_getComp2Find = "WWWWA"
    End Select
Else
    Select Case DBName
        Case "C_Code"
            CMM_getComp2Find = "ColesWeb"
        Case "C_SBCode"
            CMM_getComp2Find = "ColesSB1"
        Case "C_SBCode2"
            CMM_getComp2Find = "ColesSB2"
        Case "C_SBCode3"
            CMM_getComp2Find = "ColesSB3"
        Case "C_CCode"
            CMM_getComp2Find = "ColesPL1"
        Case "C_CCode2"
            CMM_getComp2Find = "ColesPL2"
        Case "C_CCode3"
            CMM_getComp2Find = "ColesPL3"
        Case "C_CCode4"
            CMM_getComp2Find = "ColesPL4"
        Case "C_CVCode"
            CMM_getComp2Find = "ColesVal1"
        Case "C_CVCode2"
            CMM_getComp2Find = "ColesVal2"
        Case "C_CVCode3"
            CMM_getComp2Find = "ColesVal3"
        Case "B_CBCode"
            CMM_getComp2Find = "ColesCB1"
        Case "B_PBCode"
            CMM_getComp2Find = "ColesPB1"
        Case "S_Code3"
            CMM_getComp2Find = "ColesPB2"
        Case "B_MLCode"
            CMM_getComp2Find = "ColesML1"
        Case "B_MLCode2"
            CMM_getComp2Find = "ColesML2"
        Case "B_MLCode5"
            CMM_getComp2Find = "ColesML3"
        Case "W_Code"
            CMM_getComp2Find = "WWWeb"
        Case "W_SelCode"
            CMM_getComp2Find = "WWWW1"
        Case "W_SelCode2"
            CMM_getComp2Find = "WWWW2"
        Case "W_SelCode3"
            CMM_getComp2Find = "WWWW3"
        Case "W_SelCode4"
            CMM_getComp2Find = "WWWW4"
        Case "W_SelCode5"
            CMM_getComp2Find = "WWWW5"
        Case "W_HBCode"
            CMM_getComp2Find = "WWHB1"
        Case "W_HBCode2"
            CMM_getComp2Find = "WWHB2"
        Case "W_HBCode3"
            CMM_getComp2Find = "WWHB3"
        Case "B_CBCode3"
            CMM_getComp2Find = "WWCB"
        Case "B_PBCode3"
            CMM_getComp2Find = "WWPB1"
        Case "S_Code4"
            CMM_getComp2Find = "WWPB2"
        Case "B_MLCode3"
            CMM_getComp2Find = "WWML1"
        Case "B_MLCode4"
            CMM_getComp2Find = "WWML2"
        Case "B_MLCode6"
            CMM_getComp2Find = "WWML3"
        Case "DM_Code1"
            CMM_getComp2Find = "DM1"
        Case "DM_Code2"
            CMM_getComp2Find = "DM2"
        Case "B_CBCode2"
           CMM_getComp2Find = "DMQ"
        Case "S_Code11"
            CMM_getComp2Find = "FC1"
        Case "S_Code12"
            CMM_getComp2Find = "FC2"
        Case "B_PBCode2"
            CMM_getComp2Find = "FCQ"
    End Select

End If

'If CMM_getComp2Find = "" Then
'a = a
'End If


End Function
Function CompToFindLongDesc(ByVal LongDesc As String) As String

Select Case LongDesc
    Case "Coles Control Brand": CompToFindLongDesc = "ColesCB1"
    Case "Coles Market Leader 1": CompToFindLongDesc = "ColesML1"
    Case "Coles Market Leader 2": CompToFindLongDesc = "ColesML2"
    Case "Coles Market Leader 3": CompToFindLongDesc = "ColesML3"
    Case "Coles Phantom Brand 1": CompToFindLongDesc = "ColesPB1"
    Case "Coles Phantom Brand 2": CompToFindLongDesc = "ColesPB2"
    Case "Coles Private Label 1": CompToFindLongDesc = "ColesPL1"
    Case "Coles Private Label 2": CompToFindLongDesc = "ColesPL2"
    Case "Coles Private Label 3": CompToFindLongDesc = "ColesPL3"
    Case "Coles Private Label 4": CompToFindLongDesc = "ColesPL4"
    Case "Coles Smartbuy 1": CompToFindLongDesc = "ColesSB1"
    Case "Coles Smartbuy 2": CompToFindLongDesc = "ColesSB2"
    Case "Coles Smartbuy 3": CompToFindLongDesc = "ColesSB3"
    Case "Coles Value 1": CompToFindLongDesc = "ColesVal1"
    Case "Coles Value 2": CompToFindLongDesc = "ColesVal2"
    Case "Coles Value 3": CompToFindLongDesc = "ColesVal3"
    Case "Coles Watch": CompToFindLongDesc = "ColesWeb"
    Case "Coles National Produce 1": CompToFindLongDesc = "ColesWNAT1"
    Case "Coles National Produce 2": CompToFindLongDesc = "ColesWNAT2"
    Case "Coles National Produce 3": CompToFindLongDesc = "ColesWNAT3"
    Case "Coles National Produce 4": CompToFindLongDesc = "ColesWNAT4"
    Case "Coles NSW Produce": CompToFindLongDesc = "ColesWNSW"
    Case "Coles QLD Produce": CompToFindLongDesc = "ColesWQLD"
    Case "Coles SA Produce": CompToFindLongDesc = "ColesWSA"
    Case "Coles VIC Produce": CompToFindLongDesc = "ColesWVIC"
    Case "Coles WA Produce": CompToFindLongDesc = "ColesWWA"
    Case "Dan Murphys Price 1": CompToFindLongDesc = "DM1"
    Case "Dan Murphys Price 2": CompToFindLongDesc = "DM2"
    Case "Dan Murphys Quality": CompToFindLongDesc = "DMQ"
    Case "First Choice Price 1": CompToFindLongDesc = "FC1"
    Case "First Choice Price 2": CompToFindLongDesc = "FC2"
    Case "First Choice Quality": CompToFindLongDesc = "FCQ"
    Case "Woolworths Control Brand": CompToFindLongDesc = "WWCB1"
    Case "Woolworths Homebrand 1": CompToFindLongDesc = "WWHB1"
    Case "Woolworths Homebrand 2": CompToFindLongDesc = "WWHB2"
    Case "Woolworths Homebrand 3": CompToFindLongDesc = "WWHB3"
    Case "Woolworths Market Leader 1": CompToFindLongDesc = "WWML1"
    Case "Woolworths Market Leader 2": CompToFindLongDesc = "WWML2"
    Case "Woolworths Market Leader 3": CompToFindLongDesc = "WWML3"
    Case "Woolworths Phantom Brand 1": CompToFindLongDesc = "WWPB1"
    Case "Woolworths Phantom Brand 2": CompToFindLongDesc = "WWPB2"
    Case "Woolworths Watch": CompToFindLongDesc = "WWWeb"
    Case "Woolworths National Produce 1": CompToFindLongDesc = "WWWNAT1"
    Case "Woolworths National Produce 2": CompToFindLongDesc = "WWWNAT2"
    Case "Woolworths National Produce 3": CompToFindLongDesc = "WWWNAT3"
    Case "Woolworths National Produce 4": CompToFindLongDesc = "WWWNAT4"
    Case "Woolworths NSW Produce": CompToFindLongDesc = "WWWNSW"
    Case "Woolworths QLD Produce": CompToFindLongDesc = "WWWQLD"
    Case "Woolworths SA Produce": CompToFindLongDesc = "WWWSA"
    Case "Woolworths VIC Produce": CompToFindLongDesc = "WWWVIC"
    Case "Woolworths Private Label 1": CompToFindLongDesc = "WWWW1"
    Case "Woolworths Private Label 2": CompToFindLongDesc = "WWWW2"
    Case "Woolworths Private Label 3": CompToFindLongDesc = "WWWW3"
    Case "Woolworths Private Label 4": CompToFindLongDesc = "WWWW4"
    Case "Woolworths Private Label 5": CompToFindLongDesc = "WWWW5"
    Case "Woolworths WA Produce": CompToFindLongDesc = "WWWWA"
    End Select
End Function

