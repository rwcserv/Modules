Attribute VB_Name = "CBA_PublicEnums"
                            ' CBA_Public_Enums
Public Enum e_MonthNames
    [_First]
    eJan = 1
    eFeb = 2
    eMar = 3
    eApr = 4
    eMay = 5
    eJun = 6
    eJul = 7
    eAug = 8
    eSep = 9
    eOct = 10
    eNov = 11
    eDec = 12
    [_Last]
End Enum
Public Enum e_LCG
    [_First]
    eSpirits = 1
    eSparklingwine = 2
    eWine = 3
    eBeer = 4
    eSoftDrinkAndJuices = 5
    eDetergentsAndCleaners = 6
    eHealthAndBeauty = 7
    eHygieneProducts = 8
    eBabyProducts = 9
    ePaperProducts = 10
    eWrapsAndCloths = 11
    eAudioVideoAndBatteries = 12
    eMedicine = 13
    eChildrensTextiles = 14
    eMensTextiles = 15
    eLadiesTextiles = 16
    eUnisexTextiles = 17
    eHouseholdTextiles = 18
    eFurniture = 19
    eCamerasOpticalAndEyewear = 20
    eComputersAndAccessories = 21
    eDoItYourself = 22
    eLeatherGoodsAndUmbrellas = 23
    eClocksWatchesAndJewellery = 24
    eHousewares = 25
    eFireworks = 26
    ePlantsAndFlowers = 27
    eStationeryOfficeNeedsAndGiftWrap = 28
    eDecorations = 29
    eSuitcasesAndBags = 30
    eGardening = 31
    eToys = 32
    eCarAndBike = 33
    eSportCampingAndLeisure = 34
    eHomeEntertainment = 35
    eShoes = 36
    eBooks = 37
    eFrozenFood = 38
    ePetCare = 39
    eConfectionary = 40
    eChocolates = 41
    eBiscuits = 42
    eSeasonalConfectionary = 43
    eChipsSnacksAndNuts = 44
    eCoffeeAndHotBeverages = 45
    eTea = 46
    eCannedFood = 47
    eConvenienceFoodAndSoups = 48
    eLongLifeMeats = 49
    eLongLifeDairy = 50
    eChilledFoods = 51
    eDressingOilsAndSauces = 52
    ePreservesAndSpreads = 53
    eProcessedFoods = 54
    eEggs = 55
    eRegionalBakery = 56
    eCentralBakeryAndCakes = 57
    eFruitsAndVegetables = 58
    eTobaccoProducts = 59
    ePricechanges = 60
    eStoreSupplies = 61
    eFreshmeat = 62
    eGiftcards = 63
    eFreshFish = 64
    eNewspapersAndMagazines = 65
    eLottery = 66
    [_Last]
End Enum
Public Enum e_RegionNames
    [_First]
    eNational = 500
    eMinchinbury = 501
    eDerrimut = 502
    eStapylton = 503
    ePrestons = 504
    eDandenong = 505
    eBrendale = 506
    eRegencyPark = 507
    eJadakot = 509
    [_Last]
End Enum
Public Enum e_IncoTerms
    DDP = 0
    FOB = 1
    ExWorks_FG = 2
End Enum
Public Enum e_DocuType                      ' This equates to the SH_ID / DocType / Template ID (C1_Seg_Template Hdrs in the UT_DB Access database)
    eDocuNone = 0
    eCoreRangeTEN = 1
    eRetenderTEN = 2
    eMSOTEN = 3
    eCoreRangeCategoryReview = 4                ' Camera from here???
    eProduceCategoryReview = 5
    eAlcoholCategoryReview = 6
    eSpecialsCategoryReview = 7
    eSpecSeasPerformance = 8                    ' Items 7 and 8
    eToplineCategoryPerformance = 9             ' Items 2 and 3
    eLineCountOverviewReport = 10               ' Items 10
    eCoreRangePerformance = 11                  ' Items 5,6 and 9
    eMarketOverview = 12                        ' Items 4/1 and 4/2
    eCoreRangeProductListing = 13               ' Items 20 and 21
    eForecast = 14                              ' Items 17
End Enum
Public Enum e_MSegType
    eNone = 0
    eHomescan = 1
    eScanData = 2
    eManual1 = 3
    eManual2 = 4
    eManual3 = 5
End Enum

Public Enum e_LineCountType
    eDeleted = 0
    eCore = 1
    eCurTrial = 2
    eSucTrial = 3
    eBranded = 4
    eRegional = 5
    eSeasonal = 6
    eSpecial = 7
    eBrandSucTrial = 8                          ' Enum for combination Branded and Succ Trial
End Enum
Public Enum e_CAMERAObjComp
    eDateFrom = 0
    eDateTo = 1
    eCategoryName = 2
    eLineCount = 3
    eMSegSetup = 4
    eProdToMseg = 5
End Enum
Public Enum e_LineCountPeriod ' If you update this you need to update the L0_LineCountPeriods Table in CAMERA_BE Database
    eMAT = 0
    ePriorMAT = 1
    eYTD = 2
    ePriorYTD = 3
    ePriorQTR = 4
    eQTRTD = 5
    ePriorMonth = 6
    ePriorCalendarYr = 7
    ePrevPriorCalendarYr = 8
    eCurrentMonth = 9
End Enum
Public Enum pe_CGListingColName
        eALL = 0
        eACat = 1
        eACatNum = 2
        eACGnum = 3
        eACGdesc = 4
        eASCGnum = 5
        eASCGdesc = 6
        eLegCG = 7
        eLegCGdesc = 8
        eLegSCG = 9
        eLegSCGdesc = 10
End Enum
Public Enum e_UTFldFmt        ' Refers to number of zeros in xxxxx_Ext fields i.e. UT_ID_Ext is formatted from UT_ID with eUT_ID # of zeros on the end  (UT_ID=2 >>> UT_ID_Ext=20000 (eUT_ID=4))
    [_First]                  ' These values are generated on the fly and are not kept in the database anywhere-As they are liable to exceed Long size, they are kept as strings (Mostly Keys for Dictionaries)
    eRowCol = 2               ' RowCol_Ext has eRowCol zeros
    eUT_ID = 4                ' UT_ID_Ext has eUT_ID zeros                (will be further split in Multiple(1st two #) and Group(2nd two #))
    eUT_Link = 4              ' UT_Link_ID_Ext                            If Link_ID > 0 then will format it as Link_ID and formatted zeros
    eUT_TopLeft = 5
    [_Last]
End Enum

Public Enum e_RetailorQTY      ' @RWCam Taken from cCBA_Prod & cCBA_ProdGroup - ProdGroup Retail & Qty Swapped over
    eRetail = 0
    eQTY = 1
End Enum

Public Enum e_USWorUSWpp        ' @RWCam Taken from cCBA_ProdGroup
    eUSW = 0
    eUSWp = 1
End Enum

Public Enum e_RCVMarginType     ' @RWCam Taken from cCBA_ProdGroup
    eRCVMarginPercent = 0
    eRCVContributionDollar = 1
End Enum

Public Enum e_DataCont          ' @RWCam Taken from cCBA_ProdGroup
    [_First]
    eContractData = 0
    eBraketData = 1
    eIncoData = 2
    ePricedata = 3
    [_Last]
End Enum

Public Enum e_InvMDALL          ' @RWCam Taken from cCBA_Prod
    eInventoryDifference = 0
    eMarkdowns = 1
    eBoth = 2
    eStores = 3                 ' @TP added
End Enum

Public Enum e_POSUSWTypes       ' @RWCam Taken from cCBA_Prod
    eNotUSW = 0
    eUSWisActive = 1
    eUSWALL = 2
    eProductLevel = 3
    eUSWCNT = 4
End Enum

Public Enum e_RCVQTYCostRetailNet '@TP added to allow querying of RCV
    eRCVQTY = 0
    eRCVRetail = 1
    eRCVRetailNet = 2
    eRCVCost = 3
End Enum
