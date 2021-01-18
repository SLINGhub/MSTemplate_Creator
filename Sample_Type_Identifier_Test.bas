Attribute VB_Name = "Sample_Type_Identifier_Test"
Option Explicit
Option Private Module

'@TestModule
'@Folder("Tests")

Private Assert As Object
Private Fakes As Object

'@ModuleInitialize
Public Sub ModuleInitialize()
    'this method runs once per module.
    Set Assert = CreateObject("Rubberduck.AssertClass")
    Set Fakes = CreateObject("Rubberduck.FakesProvider")
End Sub

'@ModuleCleanup
Public Sub ModuleCleanup()
    'this method runs once per module.
    Set Assert = Nothing
    Set Fakes = Nothing
End Sub

'@TestInitialize
Public Sub TestInitialize()
    'this method runs before every test in the module.
End Sub

'@TestCleanup
Public Sub TestCleanup()
    'this method runs after every test in the module.
End Sub

'@TestMethod
Public Sub isEQC_Test()
    On Error GoTo TestFail
    
    Dim EQCTestArray
    Dim i As Integer
    
    EQCTestArray = Array("EQC", "001_EQC_TQC prerun 01")
           
    For i = 0 To UBound(EQCTestArray) - LBound(EQCTestArray)
        'Debug.Print isEQC(CStr(EQCTestArray(i))) & ": " & EQCTestArray(i)
        Assert.IsTrue (isEQC(CStr(EQCTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isTQC_Test()
    On Error GoTo TestFail
    
    Dim TQCTestArray
    Dim Not_TQCTestArray
    Dim i As Integer
    
    TQCTestArray = Array("TQC", "TQC1.d", "TQC.d", "001_TQC-Eq.d", "01_TQC-1.d", "7_30m_Tqc", _
                         "TQC_PC-PE-SM_01.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_TQC_06.d", _
                         "20161117-pos-DBS-TQC-007.d", "20161117-pos-DBS-TQC-SD001-001.d", _
                         "TQCTest")
                        
    Not_TQCTestArray = Array("RQC", "010_TQCd-0", "TQC-0percent", "010_TQCd-GrpA-0", "CR_TQC-GroupB-40%", _
                             "CR_TQC-40 %", "5 percent")
           
    For i = 0 To UBound(TQCTestArray) - LBound(TQCTestArray)
        'Debug.Print isTQC(CStr(TQCTestArray(i))) & ": " & TQCTestArray(i)
        'Debug.Print (Not isRQC(CStr(TQCTestArray(i)))) & ": " & TQCTestArray(i)
        Assert.IsTrue (isTQC(CStr(TQCTestArray(i))) And Not isRQC(CStr(TQCTestArray(i))))
    Next
    
    For i = 0 To UBound(Not_TQCTestArray) - LBound(Not_TQCTestArray)
        'Debug.Print isTQC(CStr(Not_TQCTestArray(i))) & ": " & Not_TQCTestArray(i)
        'Debug.Print (Not isRQC(CStr(Not_TQCTestArray(i)))) & ": " & Not_TQCTestArray(i)
        Assert.IsFalse (isTQC(CStr(Not_TQCTestArray(i))) And Not isRQC(CStr(Not_TQCTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isBQC_Test()
    On Error GoTo TestFail
    
    Dim BQCTestArray
    Dim i As Integer
    
    BQCTestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
                         "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
                         "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
           
    For i = 0 To UBound(BQCTestArray) - LBound(BQCTestArray)
        'Debug.Print isBQC(CStr(BQCTestArray(i))) & ": " & BQCTestArray(i)
        Assert.IsTrue (isBQC(CStr(BQCTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isRQC_Test()
    On Error GoTo TestFail
    
    Dim RQCTestArray
    Dim Not_RQCTestArray
    Dim i As Integer
                            
    RQCTestArray = Array("RQC", "010_TQCd-0", "TQC-0percent", "010_TQCd-GrpA-0", "CR_TQC-GroupB-40%", _
                         "CR_TQC-40 %", "Dynamo(2)-PPG_TQCdil(040).d", "Dynamo(2)-TQCdil(050)_B.d")
                         
    Not_RQCTestArray = Array("TQC", "TQC1.d", "TQC.d", "001_TQC-Eq.d", "01_TQC-1.d", "7_30m_Tqc", _
                             "TQC_PC-PE-SM_01.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_TQC_06.d", _
                             "20161117-pos-DBS-TQC-007.d", "20161117-pos-DBS-TQC-SD001-001.d", _
                             "TQCTest", "5 percent")
           
    For i = 0 To UBound(RQCTestArray) - LBound(RQCTestArray)
        'Debug.Print isRQC(CStr(RQCTestArray(i))) & ": " & RQCTestArray(i)
        Assert.IsTrue (isRQC(CStr(RQCTestArray(i))))
    Next
    
    For i = 0 To UBound(Not_RQCTestArray) - LBound(Not_RQCTestArray)
        'Debug.Print isRQC(CStr(Not_RQCTestArray(i))) & ": " & Not_RQCTestArray(i)
        Assert.IsFalse (isRQC(CStr(Not_RQCTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isLTR_Test()
    On Error GoTo TestFail
    
    Dim LTRTestArray
    Dim i As Integer
    
    LTRTestArray = Array("LTR01.d", "LTR_01.d", "Ltr.d", "ltr.d", _
                         "018_LTR-GroupA-01")
           
    For i = 0 To UBound(LTRTestArray) - LBound(LTRTestArray)
        'Debug.Print isLTR(CStr(LTRTestArray(i))) & ": " & LTRTestArray(i)
        Assert.IsTrue (isLTR(CStr(LTRTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isNIST_Test()
    On Error GoTo TestFail
    
    Dim NISTTestArray
    Dim i As Integer
    
    NISTTestArray = Array("NIST01.d", "NIST_01.d", "Nist.d", "nist.d", _
                          "018_NIST-GroupA-01")
           
    For i = 0 To UBound(NISTTestArray) - LBound(NISTTestArray)
        'Debug.Print isNIST(CStr(NISTTestArray(i))) & ": " & NISTTestArray(i)
        Assert.IsTrue (isNIST(CStr(NISTTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isPBLK_Test()
    On Error GoTo TestFail
    
    Dim PBLKArray
    Dim i As Integer
    
    PBLKArray = Array("Blk_EXIS", _
                      "006_Extracted Blank+ISTD01", "Istd-Extract BLK", _
                      "ProcessBlank", "BlkProcessed_01", "001-PBLK")
           
    For i = 0 To UBound(PBLKArray) - LBound(PBLKArray)
        'Debug.Print isPBLK(CStr(PBLKArray(i))) & ": " & PBLKArray(i)
        Assert.IsTrue (isPBLK(CStr(PBLKArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isUBLK_Test()
    On Error GoTo TestFail
    
    Dim BlankTestArray
    Dim i As Integer
    
    BlankTestArray = Array("MyBlank", "BLK01.d", "Blk01.d", "51_blank_01.d", "07_BLANK-1.d", "001_BLK.d", "Blank_01.d", _
                           "06122016_Liver_TAG_Even01_YongLiang_blk6.d", _
                           "07062017_Marie_PL_MCF_blk_2.d", "Blk_PC-PS-SM_01.d", "Blank_1_08112016.d", _
                           "blank_7-r001.d", "Blk_12.d", "041_blank.d", "20161006-pos-SIM-Blank-001.d", _
                           "20170717-Coral lipids-dMRM-TAG-Blk-01.d", "20170623-pos-RP-blank-001.d", _
                           "20170210-RPLC-Pos-Blank02.d", "131_BLK.d")
                                 
    For i = 0 To UBound(BlankTestArray) - LBound(BlankTestArray)
        'Debug.Print (isUBLK(CStr(BlankTestArray(i)))) & ": " & BlankTestArray(i)
        'Debug.Print (Not (isPBLK(CStr(BlankTestArray(i))))) & ": " & BlankTestArray(i)
        'Debug.Print (Not (isSBLK(CStr(BlankTestArray(i))))) & ": " & BlankTestArray(i)
        Assert.IsTrue (isUBLK(CStr(BlankTestArray(i))) And Not isPBLK(CStr(BlankTestArray(i))) _
                       And Not isSBLK(CStr(BlankTestArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isSBLK_Test()
    On Error GoTo TestFail

    Dim SBLKArray
    Dim i As Integer

    SBLKArray = Array("SBLK", "Solvent_Blank", "SOL_BLK_001", _
                      "006_solvent blank", "Solvent")

    For i = 0 To UBound(SBLKArray) - LBound(SBLKArray)
        'Debug.Print isSBLK(CStr(SBLKArray(i))) & ": " & SBLKArray(i)
        Assert.IsTrue (isSBLK(CStr(SBLKArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub

'@TestMethod
Public Sub isMBLK_Test()
    On Error GoTo TestFail

    Dim MBLKArray
    Dim i As Integer

    MBLKArray = Array("MBLK", "Matrix_Blank", "MATRIX_BLK_001", _
                      "006_matrix blank", "Matrix")

    For i = 0 To UBound(MBLKArray) - LBound(MBLKArray)
        'Debug.Print isMBLK(CStr(MBLKArray(i))) & ": " & MBLKArray(i)
        Assert.IsTrue (isMBLK(CStr(MBLKArray(i))))
    Next

TestExit:
    Exit Sub
TestFail:
    Assert.Fail "Test raised an error: #" & Err.Number & " - " & Err.Description
End Sub


