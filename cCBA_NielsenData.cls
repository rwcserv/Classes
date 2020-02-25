VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCBA_NielsenData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit                     ' cCBA_NielsenData
Private pbIsHomescan As Boolean
Private plHS_ID As Long
Private plCGno As Long
Private plSCGNo As Long
Private psngRetail As Single
Private psngYOYRetail As Single
Private psngQTY As Single
Private psngYOYQTY As Single
Private psngMeasureSales As Single
Private psngYOYMeasureSales As Single
Private psngCategoryMarketShare As Single
Private psngCategoryALDIShare As Single
Private psngMarketPLShare As Single
Private psngALDIPLShare As Single
Private psngSOTRetail As Single
Private psngSOTQTY As Single
Private psngSOTMeasureSales As Single
Private pbACG As Boolean
Private psCategory As String
Private pbIsScanData As Boolean
Private pbIsManual As Boolean
Private plSD_ID As Long
Private plH_ID As Long
Private psTotGCompORALDI As String
Private psngSales As Single
Private psngSalesYOY As Single
Private psngSalesKG As Single
Private psngSalesKGYOY As Single
Private psngSalesQTY As Single
Private psngSalesQTYYOY As Single
Private psngShareSales As Single
Private psngShareSalesYOY As Single
Private psngShareSalesKG As Single
Private psngShareSalesKGYOY As Single
Private psngShareSalesQTY As Single
Private psngShareSalesQTYYOY As Single
Private plMonthNo As Long
Private plYearNo As Long
Private psMSegDescription As String
Private psSelectedForCategory As String
Private psPreviousSelectedForCategory As String
Private plCategory_ID As Long

Public Property Get IsHomescan() As Boolean: IsHomescan = pbIsHomescan: End Property
Public Property Let IsHomescan(ByVal bNewValue As Boolean): pbIsHomescan = bNewValue: End Property
Public Property Get HS_ID() As Long: HS_ID = plHS_ID: End Property
Public Property Let HS_ID(ByVal lNewValue As Long): plHS_ID = lNewValue: End Property
Public Property Get CGno() As Long: CGno = plCGno: End Property
Public Property Let CGno(ByVal lNewValue As Long): plCGno = lNewValue: End Property
Public Property Get SCGno() As Long: SCGno = plSCGNo: End Property
Public Property Let SCGno(ByVal lNewValue As Long): plSCGNo = lNewValue: End Property
Public Property Get Retail() As Single: Retail = psngRetail: End Property
Public Property Let Retail(ByVal NewValue As Single): psngRetail = NewValue: End Property
Public Property Get YOYRetail() As Single: YOYRetail = psngYOYRetail: End Property
Public Property Let YOYRetail(ByVal NewValue As Single): psngYOYRetail = NewValue: End Property
Public Property Get QTY() As Single: QTY = psngQTY: End Property
Public Property Let QTY(ByVal NewValue As Single): psngQTY = NewValue: End Property
Public Property Get YOYQTY() As Single: YOYQTY = psngYOYQTY: End Property
Public Property Let YOYQTY(ByVal NewValue As Single): psngYOYQTY = NewValue: End Property
Public Property Get MeasureSales() As Single: MeasureSales = psngMeasureSales: End Property
Public Property Let MeasureSales(ByVal NewValue As Single): psngMeasureSales = NewValue: End Property
Public Property Get YOYMeasureSales() As Single: YOYMeasureSales = psngYOYMeasureSales: End Property
Public Property Let YOYMeasureSales(ByVal NewValue As Single): psngYOYMeasureSales = NewValue: End Property
Public Property Get CategoryMarketShare() As Single: CategoryMarketShare = psngCategoryMarketShare: End Property
Public Property Let CategoryMarketShare(ByVal NewValue As Single): psngCategoryMarketShare = NewValue: End Property
Public Property Get CategoryALDIShare() As Single: CategoryALDIShare = psngCategoryALDIShare: End Property
Public Property Let CategoryALDIShare(ByVal NewValue As Single): psngCategoryALDIShare = NewValue: End Property
Public Property Get MarketPLShare() As Single: MarketPLShare = psngMarketPLShare: End Property
Public Property Let MarketPLShare(ByVal NewValue As Single): psngMarketPLShare = NewValue: End Property
Public Property Get ALDIPLShare() As Single: ALDIPLShare = psngALDIPLShare: End Property
Public Property Let ALDIPLShare(ByVal NewValue As Single): psngALDIPLShare = NewValue: End Property
Public Property Get SOTRetail() As Single: SOTRetail = psngSOTRetail: End Property
Public Property Let SOTRetail(ByVal NewValue As Single): psngSOTRetail = NewValue: End Property
Public Property Get SOTQTY() As Single: SOTQTY = psngSOTQTY: End Property
Public Property Let SOTQTY(ByVal NewValue As Single): psngSOTQTY = NewValue: End Property
Public Property Get SOTMeasureSales() As Single: SOTMeasureSales = psngSOTMeasureSales: End Property
Public Property Let SOTMeasureSales(ByVal NewValue As Single): psngSOTMeasureSales = NewValue: End Property
Public Property Get ACG() As Boolean: ACG = pbACG: End Property
Public Property Let ACG(ByVal bNewValue As Boolean): pbACG = bNewValue: End Property
Public Property Get Category() As String: Category = psCategory: End Property
Public Property Let Category(ByVal sNewValue As String): psCategory = sNewValue: End Property
Public Property Get isScanData() As Boolean: isScanData = pbIsScanData: End Property
Public Property Let isScanData(ByVal bNewValue As Boolean): pbIsScanData = bNewValue: End Property
Public Property Get SD_ID() As Long: SD_ID = plSD_ID: End Property
Public Property Let SD_ID(ByVal lNewValue As Long): plSD_ID = lNewValue: End Property
Public Property Get H_ID() As Long: H_ID = plH_ID: End Property
Public Property Let H_ID(ByVal lNewValue As Long): plH_ID = lNewValue: End Property
Public Property Get TotGComporALDI() As String: TotGComporALDI = psTotGCompORALDI: End Property
Public Property Let TotGComporALDI(ByVal sNewValue As String): psTotGCompORALDI = sNewValue: End Property
Public Property Get Sales() As Single: Sales = psngSales: End Property
Public Property Let Sales(ByVal NewValue As Single): psngSales = NewValue: End Property
Public Property Get SalesYOY() As Single: SalesYOY = psngSalesYOY: End Property
Public Property Let SalesYOY(ByVal NewValue As Single): psngSalesYOY = NewValue: End Property
Public Property Get SalesKG() As Single: SalesKG = psngSalesKG: End Property
Public Property Let SalesKG(ByVal NewValue As Single): psngSalesKG = NewValue: End Property
Public Property Get SalesKGYOY() As Single: SalesKGYOY = psngSalesKGYOY: End Property
Public Property Let SalesKGYOY(ByVal NewValue As Single): psngSalesKGYOY = NewValue: End Property
Public Property Get SalesQTY() As Single: SalesQTY = psngSalesQTY: End Property
Public Property Let SalesQTY(ByVal NewValue As Single): psngSalesQTY = NewValue: End Property
Public Property Get SalesQTYYOY() As Single: SalesQTYYOY = psngSalesQTYYOY: End Property
Public Property Let SalesQTYYOY(ByVal NewValue As Single): psngSalesQTYYOY = NewValue: End Property
Public Property Get ShareSales() As Single: ShareSales = psngShareSales: End Property
Public Property Let ShareSales(ByVal NewValue As Single): psngShareSales = NewValue: End Property
Public Property Get ShareSalesYOY() As Single: ShareSalesYOY = psngShareSalesYOY: End Property
Public Property Let ShareSalesYOY(ByVal NewValue As Single): psngShareSalesYOY = NewValue: End Property
Public Property Get ShareSalesKG() As Single: ShareSalesKG = psngShareSalesKG: End Property
Public Property Let ShareSalesKG(ByVal NewValue As Single): psngShareSalesKG = NewValue: End Property
Public Property Get ShareSalesKGYOY() As Single: ShareSalesKGYOY = psngShareSalesKGYOY: End Property
Public Property Let ShareSalesKGYOY(ByVal NewValue As Single): psngShareSalesKGYOY = NewValue: End Property
Public Property Get ShareSalesQTY() As Single: ShareSalesQTY = psngShareSalesQTY: End Property
Public Property Let ShareSalesQTY(ByVal NewValue As Single): psngShareSalesQTY = NewValue: End Property
Public Property Get ShareSalesQTYYOY() As Single: ShareSalesQTYYOY = psngShareSalesQTYYOY: End Property
Public Property Let ShareSalesQTYYOY(ByVal NewValue As Single): psngShareSalesQTYYOY = NewValue: End Property
Public Property Get MonthNo() As Long: MonthNo = plMonthNo: End Property
Public Property Let MonthNo(ByVal lNewValue As Long): plMonthNo = lNewValue: End Property
Public Property Get YearNo() As Long: YearNo = plYearNo: End Property
Public Property Let YearNo(ByVal lNewValue As Long): plYearNo = lNewValue: End Property
Public Property Get MSegDescription() As String: MSegDescription = psMSegDescription: End Property
Public Property Let MSegDescription(ByVal sNewValue As String): psMSegDescription = sNewValue: End Property
Public Property Get SelectedForCategory() As String: SelectedForCategory = psSelectedForCategory: End Property
Public Property Let SelectedForCategory(ByVal sNewValue As String): psSelectedForCategory = sNewValue: End Property
Public Property Get PreviousSelectedForCategory() As String: PreviousSelectedForCategory = psPreviousSelectedForCategory: End Property
Public Property Let PreviousSelectedForCategory(ByVal sNewValue As String): psPreviousSelectedForCategory = sNewValue: End Property
Public Property Get IsManual() As Boolean: IsManual = pbIsManual: End Property
Public Property Let IsManual(ByVal bNewValue As Boolean): pbIsManual = bNewValue: End Property
Public Property Get Category_ID() As Long: Category_ID = plCategory_ID: End Property
Public Property Let Category_ID(ByVal lNewValue As Long): plCategory_ID = lNewValue: End Property
