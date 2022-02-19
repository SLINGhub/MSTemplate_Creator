Attribute VB_Name = "Sample_Type_Identifier_Test"
Option Explicit
Option Private Module

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
'@TestMethod("Get QC Sample Type")
Public Sub Get_QC_Sample_Type_Test()
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

'' Function: isEQC_Test
'' --- Code
''  Public Sub isEQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isEQC is working
''
'' Test data are
''
''  - A string array EQCTestArray
''
'' Function will assert if Sample_Type_Identifier.isEQC gives
'' True to all entries in EQCTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isEQC_Test()
    On Error GoTo TestFail
    
    Dim EQCTestArray As Variant
    Dim arrayIndex As Integer
    
    EQCTestArray = Array("EQC", "001_EQC_TQC prerun 01")
           
    For arrayIndex = 0 To UBound(EQCTestArray) - LBound(EQCTestArray)
        'Debug.Print Sample_Type_Identifier.isEQC(CStr(EQCTestArray(arrayIndex))) & ": " & _
                     EQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isEQC(CStr(EQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isSST_Test
'' --- Code
''  Public Sub isSST_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isSST is working
''
'' Test data are
''
''  - A string array SSTTestArray
''
'' Function will assert if Sample_Type_Identifier.isSST gives
'' True to all entries in SSTTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isSST_Test()
    On Error GoTo TestFail

    Dim SSTTestArray As Variant
    Dim arrayIndex As Integer

    SSTTestArray = Array("SST01.d", "SST_01.d", "Sst.d", "sst.d", _
                     "018_SST-GroupA-01")

    For arrayIndex = 0 To UBound(SSTTestArray) - LBound(SSTTestArray)
        'Debug.Print Sample_Type_Identifier.isSST(CStr(SSTTestArray(arrayIndex))) & ": " & _
         SSTTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isSST(CStr(SSTTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isBQC_Test
'' --- Code
''  Public Sub isBQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isBQC is working
''
'' Test data are
''
''  - A string array BQCTestArray
''
'' Function will assert if Sample_Type_Identifier.isBQC gives
'' True to all entries in BQCTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isBQC_Test()
    On Error GoTo TestFail
    
    Dim BQCTestArray As Variant
    Dim arrayIndex As Integer
    
    BQCTestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
                         "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
                         "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
           
    For arrayIndex = 0 To UBound(BQCTestArray) - LBound(BQCTestArray)
        'Debug.Print Sample_Type_Identifier.isBQC(CStr(BQCTestArray(arrayIndex))) & ": " & _
                     BQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isBQC(CStr(BQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isTQC_Test
'' --- Code
''  Public Sub isTQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isTQC is working
''
'' Test data are
''
''  - A string array TQCTestArray
''  - A string array Not_TQCTestArray
''
'' Function will assert if Sample_Type_Identifier.isTQC gives
'' True to all entries in TQCTestArray and False to all entries in
'' Not_TQCTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isTQC_Test()
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
        'Debug.Print Sample_Type_Identifier.isTQC(CStr(TQCTestArray(arrayIndex))) & ": " & TQCTestArray(arrayIndex)
        'Debug.Print (Not Sample_Type_Identifier.isRQC(CStr(TQCTestArray(arrayIndex)))) & ": " & TQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isTQC(CStr(TQCTestArray(arrayIndex))) And Not Sample_Type_Identifier.isRQC(CStr(TQCTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_TQCTestArray) - LBound(Not_TQCTestArray)
        'Debug.Print Sample_Type_Identifier.isTQC(CStr(Not_TQCTestArray(arrayIndex))) & ": " & Not_TQCTestArray(arrayIndex)
        'Debug.Print (Not Sample_Type_Identifier.isRQC(CStr(Not_TQCTestArray(arrayIndex)))) & ": " & Not_TQCTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.isTQC(CStr(Not_TQCTestArray(arrayIndex))) And Not Sample_Type_Identifier.isRQC(CStr(Not_TQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isRQC_Test
'' --- Code
''  Public Sub isRQC_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isRQC is working
''
'' Test data are
''
''  - A string array RQCTestArray
''  - A string array Not_RQCTestArray
''
'' Function will assert if Sample_Type_Identifier.isRQC gives
'' True to all entries in RQCTestArray and False to all entries in
'' Not_RQCTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isRQC_Test()
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
        'Debug.Print Sample_Type_Identifier.isRQC(CStr(RQCTestArray(arrayIndex))) & ": " & _
                     RQCTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isRQC(CStr(RQCTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_RQCTestArray) - LBound(Not_RQCTestArray)
        'Debug.Print Sample_Type_Identifier.isRQC(CStr(Not_RQCTestArray(arrayIndex))) & ": " & _
                     Not_RQCTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.isRQC(CStr(Not_RQCTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isLTR_Test
'' --- Code
''  Public Sub isLTR_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isLTR is working
''
'' Test data are
''
''  - A string array LTRTestArray
''
'' Function will assert if Sample_Type_Identifier.isLTR gives
'' True to all entries in LTRTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isLTR_Test()
    On Error GoTo TestFail
    
    Dim LTRTestArray As Variant
    'Dim LTRBKTestArray As Variant
    Dim arrayIndex As Integer
    
    LTRTestArray = Array("LTR01.d", "LTR_01.d", "Ltr.d", "ltr.d", _
                         "018_LTR-GroupA-01")
                                   
    For arrayIndex = 0 To UBound(LTRTestArray) - LBound(LTRTestArray)
        'Debug.Print Sample_Type_Identifier.isLTR(CStr(LTRTestArray(arrayIndex))) & ": " & _
                     LTRTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isLTR(CStr(LTRTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isNIST_Test
'' --- Code
''  Public Sub isNIST_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isNIST is working
''
'' Test data are
''
''  - A string array NISTTestArray
''
'' Function will assert if Sample_Type_Identifier.isNIST gives
'' True to all entries in NISTTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isNIST_Test()
    On Error GoTo TestFail
    
    Dim NISTTestArray As Variant
    Dim arrayIndex As Integer
    
    NISTTestArray = Array("NIST01.d", "NIST_01.d", "Nist.d", "nist.d", _
                          "018_NIST-GroupA-01")
           
    For arrayIndex = 0 To UBound(NISTTestArray) - LBound(NISTTestArray)
        'Debug.Print Sample_Type_Identifier.isNIST(CStr(NISTTestArray(arrayIndex))) & ": " & _
        NISTTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isNIST(CStr(NISTTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isSRM_Test
'' --- Code
''  Public Sub isSRM_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isSRM is working
''
'' Test data are
''
''  - A string array SRMTestArray
''
'' Function will assert if Sample_Type_Identifier.isSRM gives
'' True to all entries in SRMTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isSRM_Test()
    On Error GoTo TestFail
    
    Dim SRMTestArray As Variant
    Dim arrayIndex As Integer
    
    SRMTestArray = Array("SRM01.d", "SRM_01.d", "Srm.d", "srm.d", _
                          "018_SRM-GroupA-01")
           
    For arrayIndex = 0 To UBound(SRMTestArray) - LBound(SRMTestArray)
        'Debug.Print Sample_Type_Identifier.isSRM(CStr(SRMTestArray(arrayIndex))) & ": " & _
                     SRMTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isSRM(CStr(SRMTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isPBLK_Test
'' --- Code
''  Public Sub isPBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isPBLK is working
''
'' Test data are
''
''  - A string array PBLKTestArray
''
'' Function will assert if Sample_Type_Identifier.isPBLK gives
'' True to all entries in PBLKTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isPBLK_Test()
    On Error GoTo TestFail
    
    Dim PBLKTestArray As Variant
    Dim arrayIndex As Integer
    
    PBLKTestArray = Array("Blk_EXIS", _
                          "006_Extracted Blank+ISTD01", "Istd-Extract BLK", _
                          "ProcessBlank", "BlkProcessed_01", "001-PBLK")
           
    For arrayIndex = 0 To UBound(PBLKTestArray) - LBound(PBLKTestArray)
        'Debug.Print Sample_Type_Identifier.isPBLK(CStr(PBLKTestArray(arrayIndex))) & ": " & _
                     PBLKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isPBLK(CStr(PBLKTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isBLK_Test
'' --- Code
''  Public Sub isBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isBLK is working
''
'' Test data are
''
''  - A string array BlankTestArray
''
'' Function will assert if Sample_Type_Identifier.isBLK gives
'' True to all entries in BlankTestArray but False when
'' Sample_Type_Identifier.isPBLK, Sample_Type_Identifier.isSBLK
'' is used instead.
''
'@TestMethod("Get QC Sample Type")
Public Sub isBLK_Test()
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
        'Debug.Print (Sample_Type_Identifier.isUBLK(CStr(BlankTestArray(arrayIndex)))) & ": " & _
                      BlankTestArray(arrayIndex)
        'Debug.Print (Not (Sample_Type_Identifier.isPBLK(CStr(BlankTestArray(arrayIndex))))) & ": " & _
                      BlankTestArray(arrayIndex)
        'Debug.Print (Not (Sample_Type_Identifier.isSBLK(CStr(BlankTestArray(arrayIndex))))) & ": " & _
                      BlankTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isBLK(CStr(BlankTestArray(arrayIndex))) And Not _
                       Sample_Type_Identifier.isPBLK(CStr(BlankTestArray(arrayIndex))) And Not _
                       Sample_Type_Identifier.isSBLK(CStr(BlankTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isSBLK_Test
'' --- Code
''  Public Sub isSBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isSBLK is working
''
'' Test data are
''
''  - A string array SBLKTestArray
''
'' Function will assert if Sample_Type_Identifier.isSBLK gives
'' True to all entries in SBLKTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isSBLK_Test()
    On Error GoTo TestFail

    Dim SBLKTestArray As Variant
    Dim arrayIndex As Integer

    SBLKTestArray = Array("SBLK", "Solvent_Blank", "SOL_BLK_001", _
                          "006_solvent blank", "Solvent")

    For arrayIndex = 0 To UBound(SBLKTestArray) - LBound(SBLKTestArray)
        'Debug.Print Sample_Type_Identifier.isSBLK(CStr(SBLKTestArray(arrayIndex))) & ": " & _
                     SBLKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isSBLK(CStr(SBLKTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isMBLK_Test
'' --- Code
''  Public Sub isMBLK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isMBLK is working
''
'' Test data are
''
''  - A string array MBLKTestArray
''
'' Function will assert if Sample_Type_Identifier.isMBLK gives
'' True to all entries in MBLKTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isMBLK_Test()
    On Error GoTo TestFail

    Dim MBLKTestArray As Variant
    Dim arrayIndex As Integer

    MBLKTestArray = Array("MBLK", "Matrix_Blank", "MATRIX_BLK_001", _
                          "006_matrix blank", "Matrix")

    For arrayIndex = 0 To UBound(MBLKTestArray) - LBound(MBLKTestArray)
        'Debug.Print Sample_Type_Identifier.isMBLK(CStr(MBLKTestArray(arrayIndex))) & ": " & _
                     MBLKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isMBLK(CStr(MBLKTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isSTD_Test
'' --- Code
''  Public Sub isSTD_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isSTD is working
''
'' Test data are
''
''  - A string array STDTestArray
''  - A string array Not_STDTestArray
''
'' Function will assert if Sample_Type_Identifier.isSTD gives
'' True to all entries in STDTestArray but False to all entries
'' in Not_STDTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isSTD_Test()
    On Error GoTo TestFail
    
    Dim STDTestArray As Variant
    Dim Not_STDTestArray As Variant
    Dim arrayIndex As Integer
    
    STDTestArray = Array("STD01.d", "STD_01.d", "Std.d", "std.d", _
                          "018_STD-GroupA-01")
                          
    Not_STDTestArray = Array("ISTD01.d", "ISTD_01.d", "Istd.d", "istd.d", _
                             "018_ISTD-GroupA-01")
           
    For arrayIndex = 0 To UBound(STDTestArray) - LBound(STDTestArray)
        'Debug.Print Sample_Type_Identifier.isSTD(CStr(STDTestArray(arrayIndex))) & ": " & _
                     STDTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isSTD(CStr(STDTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_STDTestArray) - LBound(Not_STDTestArray)
        'Debug.Print Sample_Type_Identifier.isSTD(CStr(Not_STDTestArray(arrayIndex))) & ": " & _
                     Not_STDTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.isSTD(CStr(Not_STDTestArray(arrayIndex))))
    Next

    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isLQQ_Test
'' --- Code
''  Public Sub isLQQ_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isLQQ is working
''
'' Test data are
''
''  - A string array LQQTestArray
''
'' Function will assert if Sample_Type_Identifier.isLQQ gives
'' True to all entries in LQQTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isLQQ_Test()
    On Error GoTo TestFail
    
    Dim LQQTestArray As Variant
    Dim arrayIndex As Integer
    
    LQQTestArray = Array("LQQ01.d", "LQQ_01.d", "Lqq.d", "lqq.d", _
                          "018_LQQ-GroupA-01")
                            
    For arrayIndex = 0 To UBound(LQQTestArray) - LBound(LQQTestArray)
        'Debug.Print Sample_Type_Identifier.isLQQ(CStr(LQQTestArray(arrayIndex))) & ": " & _
                     LQQTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isLQQ(CStr(LQQTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isCTRL_Test
'' --- Code
''  Public Sub isCTRL_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isCTRL is working
''
'' Test data are
''
''  - A string array CTRLTestArray
''
'' Function will assert if Sample_Type_Identifier.isCTRL gives
'' True to all entries in CTRLTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isCTRL_Test()
    On Error GoTo TestFail
    
    Dim CTRLTestArray As Variant
    Dim arrayIndex As Integer
    
    CTRLTestArray = Array("CTRL01.d", "CTRL_01.d", "Ctrl.d", "ctrl.d", _
                          "018_CTRL-GroupA-01")
                            
    For arrayIndex = 0 To UBound(CTRLTestArray) - LBound(CTRLTestArray)
        'Debug.Print Sample_Type_Identifier.isCTRL(CStr(CTRLTestArray(arrayIndex))) & ": " & _
                     CTRLTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isCTRL(CStr(CTRLTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isDUP_Test
'' --- Code
''  Public Sub isDUP_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isDUP is working
''
'' Test data are
''
''  - A string array DUPTestArray
''
'' Function will assert if Sample_Type_Identifier.isDUP gives
'' True to all entries in DUPTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isDUP_Test()
    On Error GoTo TestFail
    
    Dim DUPTestArray As Variant
    Dim arrayIndex As Integer
    
    DUPTestArray = Array("DUP01.d", "DUP_01.d", "Dup.d", "dup.d", _
                          "018_DUP-GroupA-01")
                            
    For arrayIndex = 0 To UBound(DUPTestArray) - LBound(DUPTestArray)
        'Debug.Print Sample_Type_Identifier.isDUP(CStr(DUPTestArray(arrayIndex))) & ": " & _
                     DUPTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isDUP(CStr(DUPTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isSPIK_Test
'' --- Code
''  Public Sub isSPIK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isSPIK is working
''
'' Test data are
''
''  - A string array SPIKTestArray
''
'' Function will assert if Sample_Type_Identifier.isSPIK gives
'' True to all entries in SPIKTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isSPIK_Test()
    On Error GoTo TestFail
    
    Dim SPIKTestArray As Variant
    Dim arrayIndex As Integer
    
    SPIKTestArray = Array("SPIK01.d", "SPIK_01.d", "Spik.d", _
                          "spik.d", "018_SPIK-GroupA-01", _
                          "SPIKE01.d", "SPIKE_01.d", "Spike.d", _
                          "spike.d", "018_SPIKE-GroupA-01")
                            
    For arrayIndex = 0 To UBound(SPIKTestArray) - LBound(SPIKTestArray)
        'Debug.Print Sample_Type_Identifier.isSPIK(CStr(SPIKTestArray(arrayIndex))) & ": " & _
                     SPIKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isSPIK(CStr(SPIKTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isLTRBK_Test
'' --- Code
''  Public Sub isLTRBK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isLTRBK is working
''
'' Test data are
''
''  - A string array LTRBKTestArray
''  - A string array Not_LTRBKTestArray
''
'' Function will assert if Sample_Type_Identifier.isLTRBK gives
'' True to all entries in LTRBKTestArray but False to all entries
'' in Not_LTRBKTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isLTRBK_Test()
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
        'Debug.Print Sample_Type_Identifier.isLTRBK(CStr(LTRBKTestArray(arrayIndex))) & ": " & _
                     LTRBKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isLTRBK(CStr(LTRBKTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_LTRBKTestArray) - LBound(Not_LTRBKTestArray)
        'Debug.Print Sample_Type_Identifier.isLTRBK(CStr(Not_LTRBKTestArray(arrayIndex))) & ": " & _
                     Not_LTRBKTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.isLTRBK(CStr(Not_LTRBKTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'' Function: isNISTBK_Test
'' --- Code
''  Public Sub isNISTBK_Test()
'' ---
''
'' Description:
''
'' Function used to test if the function
'' Sample_Type_Identifier.isNISTBK is working
''
'' Test data are
''
''  - A string array NISTBKTestArray
''  - A string array Not_NISTBKTestArray
''
'' Function will assert if Sample_Type_Identifier.isNISTBK gives
'' True to all entries in NISTBKTestArray but False to all entries
'' in Not_NISTBKTestArray
''
'@TestMethod("Get QC Sample Type")
Public Sub isNISTBK_Test()
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
        'Debug.Print Sample_Type_Identifier.isNISTBK(CStr(NISTBKTestArray(arrayIndex))) & ": " & _
                     NISTBKTestArray(arrayIndex)
        Assert.IsTrue (Sample_Type_Identifier.isNISTBK(CStr(NISTBKTestArray(arrayIndex))))
    Next
    
    For arrayIndex = 0 To UBound(Not_NISTBKTestArray) - LBound(Not_NISTBKTestArray)
        'Debug.Print Sample_Type_Identifier.isNISTBK(CStr(Not_NISTBKTestArray(arrayIndex))) & ": " & _
                     Not_NISTBKTestArray(arrayIndex)
        Assert.IsFalse (Sample_Type_Identifier.isNISTBK(CStr(Not_NISTBKTestArray(arrayIndex))))
    Next
    
    GoTo TestExit
TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

