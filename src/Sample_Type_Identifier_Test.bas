Attribute VB_Name = "Sample_Type_Identifier_Test"
Attribute VB_Description = "Test units for the functions in Sample_Type_Identifier Module."
Option Explicit
Option Private Module
'@ModuleDescription("Test units for the functions in Sample_Type_Identifier Module.")

'@TestModule
'@Folder("Tests")
'@IgnoreModule IntegerDataType

Private Assert As Object
'Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    'Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    'Set Fakes = Nothing
End Sub

''@TestInitialize
'Public Sub TestInitialize()
'    'this method runs before every test in the module.
'End Sub
'
''@TestCleanup
'Public Sub TestCleanup()
'    'this method runs after every test in the module.
'End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Get_QC_Sample_Type is working.")

'' Function: Get_QC_Sample_Type_Test
'' --- Code
''  Public Sub Get_QC_Sample_Type_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Get_QC_Sample_Type is working
''
'' Test data are
''
''  - A string array TestArray
''  - A string array QCArray
''
'' Function will assert if Sample_Type_Identifier.Get_QC_Sample_Type(TestArray)
'' gives the same output as QCArray
''
Public Sub Get_QC_Sample_Type_Test()
Attribute Get_QC_Sample_Type_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Get_QC_Sample_Type is working."
    On Error GoTo TestFail
    
    Dim TestArray As Variant
    Dim QCArray As Variant
    Dim arrayIndex As Integer

    TestArray = Array("EQC01", "SST02", "BQC03", "TQC04", "RQC05", _
                      "LTR06", "NIST07", "SRM08", "PBLK09", "UBLK10", _
                      "SBLK11", "MBLK12", "STD13", "LQQ14", "CTRL15", _
                      "DUP16", "SPIK17", "LTRBK18", "NISTBK19")
                      
    QCArray = Array("EQC", "SST", "BQC", "TQC", "RQC", _
                    "LTR", "NIST", "SRM", "PBLK", "UBLK", _
                    "SBLK", "MBLK", "STD", "LQQ", "CTRL", _
                    "DUP", "SPIK", "LTRBK", "NISTBK")

    For arrayIndex = 0 To UBound(TestArray) - LBound(TestArray)
        'Debug.Print Sample_Type_Identifier.Get_QC_Sample_Type(CStr(TestArray(arrayIndex))) & ": " & _
                     TestArray(arrayIndex)
        Assert.AreEqual Sample_Type_Identifier.Get_QC_Sample_Type(CStr(TestArray(arrayIndex))), _
                        CStr(QCArray(arrayIndex))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_EQC is working.")

'' Function: Is_EQC_Test
'' --- Code
''  Public Sub Is_EQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_EQC is working
''
'' Test data are
''
''  - A string array EQCTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_EQC gives
'' True to all entries in EQCTestArray
''
Public Sub Is_EQC_Test()
Attribute Is_EQC_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_EQC is working."
    On Error GoTo TestFail
    
    Dim EQCTestArray As Variant
    Dim arrayIndex As Integer
    
    EQCTestArray = Array("EQC", "001_EQC_TQC prerun 01")
           
    For arrayIndex = 0 To UBound(EQCTestArray) - LBound(EQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))) & ": " & _
                     EQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_SST is working.")

'' Function: Is_SST_Test
'' --- Code
''  Public Sub Is_SST_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_SST is working
''
'' Test data are
''
''  - A string array SSTTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_SST gives
'' True to all entries in SSTTestArray
''
Public Sub Is_SST_Test()
Attribute Is_SST_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_SST is working."
    On Error GoTo TestFail

    Dim SSTTestArray As Variant
    Dim arrayIndex As Integer

    SSTTestArray = Array("SST01.d", "SST_01.d", "Sst.d", "sst.d", _
                     "018_SST-GroupA-01")

    For arrayIndex = 0 To UBound(SSTTestArray) - LBound(SSTTestArray)
        'Debug.Print Sample_Type_Identifier.Is_SST(CStr(SSTTestArray(arrayIndex))) & ": " & _
         SSTTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_SST(CStr(SSTTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_BQC is working.")

'' Function: Is_BQC_Test
'' --- Code
''  Public Sub Is_BQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_BQC is working
''
'' Test data are
''
''  - A string array BQCTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_BQC gives
'' True to all entries in BQCTestArray
''
Public Sub Is_BQC_Test()
Attribute Is_BQC_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_BQC is working."
    On Error GoTo TestFail
    
    Dim BQCTestArray As Variant
    Dim arrayIndex As Integer
    
    BQCTestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
                         "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
                         "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
           
    For arrayIndex = 0 To UBound(BQCTestArray) - LBound(BQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_BQC(CStr(BQCTestArray(arrayIndex))) & ": " & _
                     BQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_BQC(CStr(BQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_TQC is working.")

'' Function: Is_TQC_Test
'' --- Code
''  Public Sub Is_TQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_TQC is working
''
'' Test data are
''
''  - A string array TQCTestArray
''  - A string array Not_TQCTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_TQC gives
'' True to all entries in TQCTestArray and False to all entries in
'' Not_TQCTestArray
''
Public Sub Is_TQC_Test()
Attribute Is_TQC_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_TQC is working."
    On Error GoTo TestFail
    
    Dim TQCTestArray As Variant
    Dim Not_TQCTestArray As Variant
    Dim arrayIndex As Integer
    
    TQCTestArray = Array("TQC", "TQC1.d", "TQC.d", "001_TQC-Eq.d", "01_TQC-1.d", "7_30m_Tqc", _
                         "TQC_PC-PE-SM_01.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_TQC_06.d", _
                         "20161117-pos-DBS-TQC-007.d", "20161117-pos-DBS-TQC-SD001-001.d", _
                         "TQCTest")
                        
    Not_TQCTestArray = Array("RQC", "010_TQCd-0", "TQC-0percent", "010_TQCd-GrpA-0", "CR_TQC-GroupB-40%", _
                             "CR_TQC-40 %", "5 percent")
           
    For arrayIndex = 0 To UBound(TQCTestArray) - LBound(TQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_TQC(CStr(TQCTestArray(arrayIndex))) & ": " & TQCTestArray(arrayIndex)
        'Debug.Print (Not Sample_Type_Identifier.Is_RQC(CStr(TQCTestArray(arrayIndex)))) & ": " & TQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_TQC(CStr(TQCTestArray(arrayIndex))) And Not Sample_Type_Identifier.Is_RQC(CStr(TQCTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_TQCTestArray) - LBound(Not_TQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_TQC(CStr(Not_TQCTestArray(arrayIndex))) & ": " & Not_TQCTestArray(arrayIndex)
        'Debug.Print (Not Sample_Type_Identifier.Is_RQC(CStr(Not_TQCTestArray(arrayIndex)))) & ": " & Not_TQCTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.Is_TQC(CStr(Not_TQCTestArray(arrayIndex))) And Not Sample_Type_Identifier.Is_RQC(CStr(Not_TQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_RQC is working.")

'' Function: Is_RQC_Test
'' --- Code
''  Public Sub Is_RQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_RQC is working
''
'' Test data are
''
''  - A string array RQCTestArray
''  - A string array Not_RQCTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_RQC gives
'' True to all entries in RQCTestArray and False to all entries in
'' Not_RQCTestArray
''
Public Sub Is_RQC_Test()
Attribute Is_RQC_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_RQC is working."
    On Error GoTo TestFail
    
    Dim RQCTestArray As Variant
    Dim Not_RQCTestArray As Variant
    Dim arrayIndex As Integer
                            
    RQCTestArray = Array("RQC", "010_TQCd-0", "TQC-0percent", "010_TQCd-GrpA-0", "CR_TQC-GroupB-40%", _
                         "CR_TQC-40 %", "Dynamo(2)-PPG_TQCdil(040).d", "Dynamo(2)-TQCdil(050)_B.d")
                         
    Not_RQCTestArray = Array("TQC", "TQC1.d", "TQC.d", "001_TQC-Eq.d", "01_TQC-1.d", "7_30m_Tqc", _
                             "TQC_PC-PE-SM_01.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_TQC_06.d", _
                             "20161117-pos-DBS-TQC-007.d", "20161117-pos-DBS-TQC-SD001-001.d", _
                             "TQCTest", "5 percent")
           
    For arrayIndex = 0 To UBound(RQCTestArray) - LBound(RQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_RQC(CStr(RQCTestArray(arrayIndex))) & ": " & _
                     RQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_RQC(CStr(RQCTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_RQCTestArray) - LBound(Not_RQCTestArray)
        'Debug.Print Sample_Type_Identifier.Is_RQC(CStr(Not_RQCTestArray(arrayIndex))) & ": " & _
                     Not_RQCTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.Is_RQC(CStr(Not_RQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_LTR is working.")

'' Function: Is_LTR_Test
'' --- Code
''  Public Sub Is_LTR_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_LTR is working
''
'' Test data are
''
''  - A string array LTRTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_LTR gives
'' True to all entries in LTRTestArray
''
Public Sub Is_LTR_Test()
Attribute Is_LTR_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_LTR is working."
    On Error GoTo TestFail
    
    Dim LTRTestArray As Variant
    'Dim LTRBKTestArray As Variant
    Dim arrayIndex As Integer
    
    LTRTestArray = Array("LTR01.d", "LTR_01.d", "Ltr.d", "ltr.d", _
                         "018_LTR-GroupA-01")
                                   
    For arrayIndex = 0 To UBound(LTRTestArray) - LBound(LTRTestArray)
        'Debug.Print Sample_Type_Identifier.Is_LTR(CStr(LTRTestArray(arrayIndex))) & ": " & _
                     LTRTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_LTR(CStr(LTRTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_NIST is working.")

'' Function: Is_NIST_Test
'' --- Code
''  Public Sub Is_NIST_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_NIST is working
''
'' Test data are
''
''  - A string array NISTTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_NIST gives
'' True to all entries in NISTTestArray
''
Public Sub Is_NIST_Test()
Attribute Is_NIST_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_NIST is working."
    On Error GoTo TestFail
    
    Dim NISTTestArray As Variant
    Dim arrayIndex As Integer
    
    NISTTestArray = Array("NIST01.d", "NIST_01.d", "Nist.d", "nist.d", _
                          "018_NIST-GroupA-01")
           
    For arrayIndex = 0 To UBound(NISTTestArray) - LBound(NISTTestArray)
        'Debug.Print Sample_Type_Identifier.Is_NIST(CStr(NISTTestArray(arrayIndex))) & ": " & _
        NISTTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_NIST(CStr(NISTTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_SRM is working.")

'' Function: Is_SRM_Test
'' --- Code
''  Public Sub Is_SRM_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_SRM is working
''
'' Test data are
''
''  - A string array SRMTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_SRM gives
'' True to all entries in SRMTestArray
''
Public Sub Is_SRM_Test()
Attribute Is_SRM_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_SRM is working."
    On Error GoTo TestFail
    
    Dim SRMTestArray As Variant
    Dim arrayIndex As Integer
    
    SRMTestArray = Array("SRM01.d", "SRM_01.d", "Srm.d", "srm.d", _
                          "018_SRM-GroupA-01")
           
    For arrayIndex = 0 To UBound(SRMTestArray) - LBound(SRMTestArray)
        'Debug.Print Sample_Type_Identifier.Is_SRM(CStr(SRMTestArray(arrayIndex))) & ": " & _
                     SRMTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_SRM(CStr(SRMTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_SRM is working.")

'' Function: Is_PBLK_Test
'' --- Code
''  Public Sub Is_PBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_PBLK is working
''
'' Test data are
''
''  - A string array PBLKTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_PBLK gives
'' True to all entries in PBLKTestArray
''
Public Sub Is_PBLK_Test()
Attribute Is_PBLK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_SRM is working."
    On Error GoTo TestFail
    
    Dim PBLKTestArray As Variant
    Dim arrayIndex As Integer
    
    PBLKTestArray = Array("Blk_EXIS", _
                          "006_Extracted Blank+ISTD01", "Istd-Extract BLK", _
                          "ProcessBlank", "BlkProcessed_01", "001-PBLK")
           
    For arrayIndex = 0 To UBound(PBLKTestArray) - LBound(PBLKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_PBLK(CStr(PBLKTestArray(arrayIndex))) & ": " & _
                     PBLKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_PBLK(CStr(PBLKTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_BLK is working.")

'' Function: Is_BLK_Test
'' --- Code
''  Public Sub Is_BLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_BLK is working
''
'' Test data are
''
''  - A string array BlankTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_BLK gives
'' True to all entries in BlankTestArray but False when
'' Sample_Type_Identifier.Is_PBLK, Sample_Type_Identifier.Is_SBLK
'' is used instead.
''
Public Sub Is_BLK_Test()
Attribute Is_BLK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_BLK is working."
    On Error GoTo TestFail
    
    Dim BlankTestArray As Variant
    Dim arrayIndex As Integer
    
    BlankTestArray = Array("MyBlank", "BLK01.d", "Blk01.d", "51_blank_01.d", "07_BLANK-1.d", "001_BLK.d", "Blank_01.d", _
                           "06122016_Liver_TAG_Even01_YongLiang_blk6.d", _
                           "07062017_Marie_PL_MCF_blk_2.d", "Blk_PC-PS-SM_01.d", "Blank_1_08112016.d", _
                           "blank_7-r001.d", "Blk_12.d", "041_blank.d", "20161006-pos-SIM-Blank-001.d", _
                           "20170717-Coral lipids-dMRM-TAG-Blk-01.d", "20170623-pos-RP-blank-001.d", _
                           "20170210-RPLC-Pos-Blank02.d", "131_BLK.d")
                                 
    For arrayIndex = 0 To UBound(BlankTestArray) - LBound(BlankTestArray)
        'Debug.Print (Sample_Type_Identifier.Is_UBLK(CStr(BlankTestArray(arrayIndex)))) & ": " & _
                      BlankTestArray(arrayIndex)
        'Debug.Print (Not (Sample_Type_Identifier.Is_PBLK(CStr(BlankTestArray(arrayIndex))))) & ": " & _
                      BlankTestArray(arrayIndex)
        'Debug.Print (Not (Sample_Type_Identifier.Is_SBLK(CStr(BlankTestArray(arrayIndex))))) & ": " & _
                      BlankTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_BLK(CStr(BlankTestArray(arrayIndex))) And Not _
                       Sample_Type_Identifier.Is_PBLK(CStr(BlankTestArray(arrayIndex))) And Not _
                       Sample_Type_Identifier.Is_SBLK(CStr(BlankTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_SBLK is working.")

'' Function: Is_SBLK_Test
'' --- Code
''  Public Sub Is_SBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_SBLK is working
''
'' Test data are
''
''  - A string array SBLKTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_SBLK gives
'' True to all entries in SBLKTestArray
''
Public Sub Is_SBLK_Test()
Attribute Is_SBLK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_SBLK is working."
    On Error GoTo TestFail

    Dim SBLKTestArray As Variant
    Dim arrayIndex As Integer

    SBLKTestArray = Array("SBLK", "Solvent_Blank", "SOL_BLK_001", _
                          "006_solvent blank", "Solvent")

    For arrayIndex = 0 To UBound(SBLKTestArray) - LBound(SBLKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_SBLK(CStr(SBLKTestArray(arrayIndex))) & ": " & _
                     SBLKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_SBLK(CStr(SBLKTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_MBLK is working.")

'' Function: Is_MBLK_Test
'' --- Code
''  Public Sub Is_MBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_MBLK is working
''
'' Test data are
''
''  - A string array MBLKTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_MBLK gives
'' True to all entries in MBLKTestArray
''
Public Sub Is_MBLK_Test()
Attribute Is_MBLK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_MBLK is working."
    On Error GoTo TestFail

    Dim MBLKTestArray As Variant
    Dim arrayIndex As Integer

    MBLKTestArray = Array("MBLK", "Matrix_Blank", "MATRIX_BLK_001", _
                          "006_matrix blank", "Matrix")

    For arrayIndex = 0 To UBound(MBLKTestArray) - LBound(MBLKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_MBLK(CStr(MBLKTestArray(arrayIndex))) & ": " & _
                     MBLKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_MBLK(CStr(MBLKTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_STD is working.")

'' Function: Is_STD_Test
'' --- Code
''  Public Sub Is_STD_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_STD is working
''
'' Test data are
''
''  - A string array STDTestArray
''  - A string array Not_STDTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_STD gives
'' True to all entries in STDTestArray but False to all entries
'' in Not_STDTestArray
''
Public Sub Is_STD_Test()
Attribute Is_STD_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_STD is working."
    On Error GoTo TestFail
    
    Dim STDTestArray As Variant
    Dim Not_STDTestArray As Variant
    Dim arrayIndex As Integer
    
    STDTestArray = Array("STD01.d", "STD_01.d", "Std.d", "std.d", _
                          "018_STD-GroupA-01")
                          
    Not_STDTestArray = Array("ISTD01.d", "ISTD_01.d", "Istd.d", "istd.d", _
                             "018_ISTD-GroupA-01")
           
    For arrayIndex = 0 To UBound(STDTestArray) - LBound(STDTestArray)
        'Debug.Print Sample_Type_Identifier.Is_STD(CStr(STDTestArray(arrayIndex))) & ": " & _
                     STDTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_STD(CStr(STDTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_STDTestArray) - LBound(Not_STDTestArray)
        'Debug.Print Sample_Type_Identifier.Is_STD(CStr(Not_STDTestArray(arrayIndex))) & ": " & _
                     Not_STDTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.Is_STD(CStr(Not_STDTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_LQQ is working.")

'' Function: Is_LQQ_Test
'' --- Code
''  Public Sub Is_LQQ_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_LQQ is working
''
'' Test data are
''
''  - A string array LQQTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_LQQ gives
'' True to all entries in LQQTestArray
''
Public Sub Is_LQQ_Test()
Attribute Is_LQQ_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_LQQ is working."
    On Error GoTo TestFail
    
    Dim LQQTestArray As Variant
    Dim arrayIndex As Integer
    
    LQQTestArray = Array("LQQ01.d", "LQQ_01.d", "Lqq.d", "lqq.d", _
                          "018_LQQ-GroupA-01")
                            
    For arrayIndex = 0 To UBound(LQQTestArray) - LBound(LQQTestArray)
        'Debug.Print Sample_Type_Identifier.Is_LQQ(CStr(LQQTestArray(arrayIndex))) & ": " & _
                     LQQTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_LQQ(CStr(LQQTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_CTRL is working.")

'' Function: Is_CTRL_Test
'' --- Code
''  Public Sub Is_CTRL_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_CTRL is working
''
'' Test data are
''
''  - A string array CTRLTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_CTRL gives
'' True to all entries in CTRLTestArray
''
Public Sub Is_CTRL_Test()
Attribute Is_CTRL_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_CTRL is working."
    On Error GoTo TestFail
    
    Dim CTRLTestArray As Variant
    Dim arrayIndex As Integer
    
    CTRLTestArray = Array("CTRL01.d", "CTRL_01.d", "Ctrl.d", "ctrl.d", _
                          "018_CTRL-GroupA-01")
                            
    For arrayIndex = 0 To UBound(CTRLTestArray) - LBound(CTRLTestArray)
        'Debug.Print Sample_Type_Identifier.Is_CTRL(CStr(CTRLTestArray(arrayIndex))) & ": " & _
                     CTRLTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_CTRL(CStr(CTRLTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_DUP is working.")

'' Function: Is_DUP_Test
'' --- Code
''  Public Sub Is_DUP_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_DUP is working
''
'' Test data are
''
''  - A string array DUPTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_DUP gives
'' True to all entries in DUPTestArray
''
Public Sub Is_DUP_Test()
Attribute Is_DUP_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_DUP is working."
    On Error GoTo TestFail
    
    Dim DUPTestArray As Variant
    Dim arrayIndex As Integer
    
    DUPTestArray = Array("DUP01.d", "DUP_01.d", "Dup.d", "dup.d", _
                          "018_DUP-GroupA-01")
                            
    For arrayIndex = 0 To UBound(DUPTestArray) - LBound(DUPTestArray)
        'Debug.Print Sample_Type_Identifier.Is_DUP(CStr(DUPTestArray(arrayIndex))) & ": " & _
                     DUPTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_DUP(CStr(DUPTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_SPIK is working.")

'' Function: Is_SPIK_Test
'' --- Code
''  Public Sub Is_SPIK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_SPIK is working
''
'' Test data are
''
''  - A string array SPIKTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_SPIK gives
'' True to all entries in SPIKTestArray
''
Public Sub Is_SPIK_Test()
Attribute Is_SPIK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_SPIK is working."
    On Error GoTo TestFail
    
    Dim SPIKTestArray As Variant
    Dim arrayIndex As Integer
    
    SPIKTestArray = Array("SPIK01.d", "SPIK_01.d", "Spik.d", _
                          "spik.d", "018_SPIK-GroupA-01", _
                          "SPIKE01.d", "SPIKE_01.d", "Spike.d", _
                          "spike.d", "018_SPIKE-GroupA-01")
                            
    For arrayIndex = 0 To UBound(SPIKTestArray) - LBound(SPIKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_SPIK(CStr(SPIKTestArray(arrayIndex))) & ": " & _
                     SPIKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_SPIK(CStr(SPIKTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_LTRBK is working.")

'' Function: Is_LTRBK_Test
'' --- Code
''  Public Sub Is_LTRBK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_LTRBK is working
''
'' Test data are
''
''  - A string array LTRBKTestArray
''  - A string array Not_LTRBKTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_LTRBK gives
'' True to all entries in LTRBKTestArray but False to all entries
'' in Not_LTRBKTestArray
''
Public Sub Is_LTRBK_Test()
Attribute Is_LTRBK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_LTRBK is working."
    On Error GoTo TestFail
    
    Dim LTRBKTestArray As Variant
    Dim Not_LTRBKTestArray As Variant
    Dim arrayIndex As Integer
    
    LTRBKTestArray = Array("LTRBK01.d", "LTRBK_01.d", "Ltrbk.d", _
                           "ltrbk.d", "018_LTRBK-GroupA-01", _
                           "LTRBLK01.d", "LTRBLK_01.d", "Ltrblk.d", _
                           "ltrblk.d", "018_LTRBLK-GroupA-01", _
                           "LTR_BLK01.d", "LTR_BLK_01.d", "Ltr_blk.d", _
                           "ltr_blk.d", "018_LTR_BLK-GroupA-01", _
                           "LTR-BLK01.d", "LTR-BLK_01.d", "Ltr-blk.d", _
                           "ltr-blk.d", "018_LTR-BLK-GroupA-01", _
                           "LTR BLK01.d", "LTR BLK_01.d", "Ltr blk.d", _
                           "ltr blk.d", "018_LTR BLK-GroupA-01", _
                           "No_ISTD_LTR.d", "NoIS_LTR.d", _
                           "LTR_NoIS.d", "LTR-no_istd.d")
    Not_LTRBKTestArray = Array("LTR01.d", "LTR_01.d", "Ltr.d", "ltr.d", _
                               "018_LTR-GroupA-01", "NO_LTR.d")
                            
    For arrayIndex = 0 To UBound(LTRBKTestArray) - LBound(LTRBKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_LTRBK(CStr(LTRBKTestArray(arrayIndex))) & ": " & _
                     LTRBKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_LTRBK(CStr(LTRBKTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_LTRBKTestArray) - LBound(Not_LTRBKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_LTRBK(CStr(Not_LTRBKTestArray(arrayIndex))) & ": " & _
                     Not_LTRBKTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.Is_LTRBK(CStr(Not_LTRBKTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod("Get QC Sample Type")
'@Description("Function used to test if the function Sample_Type_Identifier.Is_NISTBK is working.")

'' Function: Is_NISTBK_Test
'' --- Code
''  Public Sub Is_NISTBK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.Is_NISTBK is working
''
'' Test data are
''
''  - A string array NISTBKTestArray
''  - A string array Not_NISTBKTestArray
''
'' Function will assert if Sample_Type_Identifier.Is_NISTBK gives
'' True to all entries in NISTBKTestArray but False to all entries
'' in Not_NISTBKTestArray
''
Public Sub Is_NISTBK_Test()
Attribute Is_NISTBK_Test.VB_Description = "Function used to test if the function Sample_Type_Identifier.Is_NISTBK is working."
    On Error GoTo TestFail
    
    Dim NISTBKTestArray As Variant
    Dim Not_NISTBKTestArray As Variant
    Dim arrayIndex As Integer
    
    NISTBKTestArray = Array("NISTBK01.d", "NISTBK_01.d", "Nistbk.d", _
                            "nistbk.d", "018_NISTBK-GroupA-01", _
                            "NISTBLK01.d", "NISTBLK_01.d", "Nistblk.d", _
                            "nistblk.d", "018_NISTBLK-GroupA-01", _
                            "NIST_BLK01.d", "NIST_BLK_01.d", "Nist_blk.d", _
                            "nist_blk.d", "018_NIST_BLK-GroupA-01", _
                            "NIST-BLK01.d", "NIST-BLK_01.d", "Nist-blk.d", _
                            "nist-blk.d", "018_NIST-BLK-GroupA-01", _
                            "NIST BLK01.d", "NIST BLK_01.d", "Nist blk.d", _
                            "nist blk.d", "018_NIST BLK-GroupA-01", _
                            "No_ISTD_NIST.d", "NoIS_NIST.d", _
                            "NIST_NoIS.d", "NIST-no_istd.d")
    Not_NISTBKTestArray = Array("NIST01.d", "NIST_01.d", "Nist.d", "nist.d", _
                                "018_NIST-GroupA-01", "NO_NIST.d")
                            
    For arrayIndex = 0 To UBound(NISTBKTestArray) - LBound(NISTBKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_NISTBK(CStr(NISTBKTestArray(arrayIndex))) & ": " & _
                     NISTBKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.Is_NISTBK(CStr(NISTBKTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_NISTBKTestArray) - LBound(Not_NISTBKTestArray)
        'Debug.Print Sample_Type_Identifier.Is_NISTBK(CStr(Not_NISTBKTestArray(arrayIndex))) & ": " & _
                     Not_NISTBKTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.Is_NISTBK(CStr(Not_NISTBKTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub
