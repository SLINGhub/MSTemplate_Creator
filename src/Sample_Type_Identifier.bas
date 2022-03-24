Attribute VB_Name = "Sample_Type_Identifier"
Option Explicit
'@Folder("Sample Annot Functions")

'' Function: Get_QC_Sample_Type
''
'' Description:
''
'' Get the sample type based on the input string (sample name).
''
'' Parameters:
''
''    FileName - Input string to check what sample type it belongs to
''
'' Returns:
''    A string indicated the sample type the sample belongs to
''
'' Examples:
''
'' --- Code
''    Dim TestArray As Variant
''    Dim arrayIndex As Integer
''
''    TestArray = Array("EQC01", "SST02", "BQC03", "TQC04", "RQC05", _
''                      "LTR06", "NIST07", "SRM08", "PBLK09", "UBLK10", _
''                      "SBLK11", "MBLK12", "STD13", "LQQ14", "CTRL15", _
''                      "DUP16", "SPIK17", "LTRBK18", "NISTBK19")
''
''    For arrayIndex = 0 To UBound(TestArray) - LBound(TestArray)
''        Debug.Print Sample_Type_Identifier.Get_QC_Sample_Type(CStr(TestArray(arrayIndex))) & ": " & _
''                    TestArray(arrayIndex)
''    Next
'' ---
Public Function Get_QC_Sample_Type(ByVal FileName As String) As String
    If Is_EQC(FileName) Then
        Get_QC_Sample_Type = "EQC"
    ElseIf Is_SST(FileName) Then
        Get_QC_Sample_Type = "SST"
    ElseIf Is_BQC(FileName) Then
        Get_QC_Sample_Type = "BQC"
    ElseIf Is_TQC(FileName) And Not Is_RQC(FileName) Then
        Get_QC_Sample_Type = "TQC"
    ElseIf Is_RQC(FileName) Then
        Get_QC_Sample_Type = "RQC"
    ElseIf Is_LTR(FileName) And Not Is_LTRBK(FileName) Then
        Get_QC_Sample_Type = "LTR"
    ElseIf Is_NIST(FileName) And Not Is_NISTBK(FileName) Then
        Get_QC_Sample_Type = "NIST"
    ElseIf Is_SRM(FileName) Then
        Get_QC_Sample_Type = "SRM"
    ElseIf Is_PBLK(FileName) Then
        Get_QC_Sample_Type = "PBLK"
    ElseIf Is_BLK(FileName) And Not Is_PBLK(FileName) And Not Is_SBLK(FileName) _
        And Not Is_MBLK(FileName) And Not Is_LTRBK(FileName) _
        And Not Is_NISTBK(FileName) Then
        Get_QC_Sample_Type = "UBLK"
    ElseIf Is_SBLK(FileName) Then
        Get_QC_Sample_Type = "SBLK"
    ElseIf Is_MBLK(FileName) Then
        Get_QC_Sample_Type = "MBLK"
    ElseIf Is_STD(FileName) Then
        Get_QC_Sample_Type = "STD"
    ElseIf Is_LQQ(FileName) Then
        Get_QC_Sample_Type = "LQQ"
    ElseIf Is_CTRL(FileName) Then
        Get_QC_Sample_Type = "CTRL"
    ElseIf Is_DUP(FileName) Then
        Get_QC_Sample_Type = "DUP"
    ElseIf Is_SPIK(FileName) Then
        Get_QC_Sample_Type = "SPIK"
    ElseIf Is_LTRBK(FileName) Then
        Get_QC_Sample_Type = "LTRBK"
    ElseIf Is_NISTBK(FileName) Then
        Get_QC_Sample_Type = "NISTBK"
    End If
    
End Function

'' Function: Is_EQC
''
'' Description:
''
'' Check is the input string (sample name) is an EQC.
''
'' Parameters:
''
''    FileName - Input string to check if it is an EQC
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "EQC", "Eqc", "eqc"
''
'' Examples:
''
'' --- Code
''    Dim EQCTestArray As Variant
''    Dim arrayIndex As Integer
''
''    EQCTestArray = Array("EQC", "001_EQC_TQC prerun 01")
''
''    For arrayIndex = 0 To UBound(EQCTestArray) - LBound(EQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_EQC(CStr(EQCTestArray(arrayIndex))) & ": " & _
''                    EQCTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_EQC(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim EQCRegEx As RegExp
    Set EQCRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    EQCRegEx.Pattern = "(EQC|[Ee]qc)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_EQC = EQCRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_SST
''
'' Description:
''
'' Check is the input string (sample name) is a SST.
''
'' Parameters:
''
''    FileName - Input string to check if it is a SST
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "SST", "Stt", "stt"
''
'' Examples:
''
'' --- Code
''    Dim SSTArray As Variant
''    Dim arrayIndex As Integer
''
''    SSTTestArray = Array("SST01.d", "SST_01.d", "Sst.d", "sst.d", _
''                     "018_SST-GroupA-01")
''
''    For arrayIndex = 0 To UBound(SSTTestArray) - LBound(SSTTestArray)
''        Debug.Print Sample_Type_Identifier.Is_SST(CStr(SSTTestArray(arrayIndex))) & ": " & _
''                    SSTTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_SST(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim SSTRegEx As RegExp
    Set SSTRegEx = New RegExp
    
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    SSTRegEx.Pattern = "(SST|[Ss]st)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_SST = SSTRegEx.Test(OnlyLettersText)
End Function

'' Function: Is_BQC
''
'' Description:
''
'' Check is the input string (sample name) is a BQC.
''
'' Parameters:
''
''    FileName - Input string to check if it is a BQC
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "BQC", "Bqc", "bqc",
''    "PQC", "Pqc", "pqc"
''
'' Examples:
''
'' --- Code
''    Dim BQCTestArray As Variant
''    Dim arrayIndex As Integer
''
''    BQCTestArray = Array("11_PQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_PQC_2.d", _
''                         "PQC_PC-PE-SM_01.d", "PQC1_LPC-LPE-PG-PI-PS_01.d", "PQC1_02.d", _
''                         "11_BQC-2.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_BQC_2.d", "018_BQC_PQC01")
''
''    For arrayIndex = 0 To UBound(BQCTestArray) - LBound(BQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_BQC(CStr(BQCTestArray(arrayIndex))) & ": " & _
''                    BQCTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_BQC(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim BQCRegEx As RegExp
    Set BQCRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    BQCRegEx.Pattern = "([BP]QC|[Pp]qc|[Bb]qc)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_BQC = BQCRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_TQC
''
'' Description:
''
'' Check is the input string (sample name) is a TQC.
''
'' Parameters:
''
''    FileName - Input string to check if it is a TQC
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "TQC", "Tqc", "tqc"
''
'' Examples:
''
'' --- Code
''    Dim TQCTestArray As Variant
''    Dim Not_TQCTestArray As Variant
''    Dim arrayIndex As Integer
''
''    TQCTestArray = Array("TQC", "TQC1.d", "TQC.d", "001_TQC-Eq.d", "01_TQC-1.d", "7_30m_Tqc", _
''                         "TQC_PC-PE-SM_01.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_TQC_06.d", _
''                         "20161117-pos-DBS-TQC-007.d", "20161117-pos-DBS-TQC-SD001-001.d", _
''                         "TQCTest")
''
''    Not_TQCTestArray = Array("RQC", "010_TQCd-0", "TQC-0percent", "010_TQCd-GrpA-0", "CR_TQC-GroupB-40%", _
''                             "CR_TQC-40 %", "5 percent")
''
''    For arrayIndex = 0 To UBound(TQCTestArray) - LBound(TQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_TQC(CStr(TQCTestArray(arrayIndex))) & ": " & TQCTestArray(arrayIndex)
''        Debug.Print (Not Sample_Type_Identifier.Is_RQC(CStr(TQCTestArray(arrayIndex)))) & ": " & TQCTestArray(arrayIndex)
''    Next
''
''    For arrayIndex = 0 To UBound(Not_TQCTestArray) - LBound(Not_TQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_TQC(CStr(Not_TQCTestArray(arrayIndex))) & ": " & Not_TQCTestArray(arrayIndex)
''        Debug.Print (Not Sample_Type_Identifier.Is_RQC(CStr(Not_TQCTestArray(arrayIndex)))) & ": " & Not_TQCTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_TQC(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim TQCRegEx As RegExp
    Set TQCRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    TQCRegEx.Pattern = "(TQC|[Tt]qc)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_TQC = TQCRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_RQC
''
'' Description:
''
'' Check is the input string (sample name) is a RQC.
''
'' Parameters:
''
''    FileName - Input string to check if it is a RQC
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "RQC", "Rqc", "rqc",
''    "TQCd", "TQCdil", "TQC" with percent, "TQC" with %.
''
'' Examples:
''
'' --- Code
''    Dim RQCTestArray As Variant
''    Dim Not_RQCTestArray As Variant
''    Dim arrayIndex As Integer
''
''    RQCTestArray = Array("RQC", "010_TQCd-0", "TQC-0percent", "010_TQCd-GrpA-0", "CR_TQC-GroupB-40%", _
''                         "CR_TQC-40 %", "Dynamo(2)-PPG_TQCdil(040).d", "Dynamo(2)-TQCdil(050)_B.d")
''
''    Not_RQCTestArray = Array("TQC", "TQC1.d", "TQC.d", "001_TQC-Eq.d", "01_TQC-1.d", "7_30m_Tqc", _
''                             "TQC_PC-PE-SM_01.d", "20032017_TAG_SNEHAMTHD__Dogs_PL_TQC_06.d", _
''                             "20161117-pos-DBS-TQC-007.d", "20161117-pos-DBS-TQC-SD001-001.d", _
''                             "TQCTest", "5 percent")
''
''    For arrayIndex = 0 To UBound(RQCTestArray) - LBound(RQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_RQC(CStr(RQCTestArray(arrayIndex))) & ": " & _
''                    RQCTestArray(arrayIndex)
''    Next
''
''    For arrayIndex = 0 To UBound(Not_RQCTestArray) - LBound(Not_RQCTestArray)
''        Debug.Print Sample_Type_Identifier.Is_RQC(CStr(Not_RQCTestArray(arrayIndex))) & ": " & _
''                    Not_RQCTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_RQC(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim TQCdRegEx As RegExp
    Set TQCdRegEx = New RegExp
    Dim TQCno_dRegEx As RegExp
    Set TQCno_dRegEx = New RegExp
    Dim RQCRegEx As RegExp
    Set RQCRegEx = New RegExp
    'Dim PercentRegEx As RegExp
    'Set PercentRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    'If there is a d, percent is not compulsory
    TQCdRegEx.Pattern = "(TQCd(il)?|[Tt]qcd(il)?)" & "([\s,_,-]+)?" & _
                        "([A-Za-z0-9]+)?" & "([\s,_,-]+)?" & _
                        "((\()?[0-9]+(\))?)" & "([\s,_,-]+)?" & _
                        "([A-Za-z0-9]+)?" & "([\s,_,-]+)?" & "([Pp]ercent|%)?"
    'If no d, percent is compulsory
    TQCno_dRegEx.Pattern = "(TQC|[Tt]qc)" & "([\s,_,-]+)?" & _
                           "([A-Za-z0-9]+)?" & "([\s,_,-]+)?" & _
                           "((\()?[0-9]+(\))?)" & "([\s,_,-]+)?" & _
                           "([A-Za-z0-9]+)?" & "([\s,_,-]+)?" & "([Pp]ercent|%)"
    
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    RQCRegEx.Pattern = "(RQC|[Rr]qc)"
    
    Is_RQC = TQCdRegEx.Test(FileName) Or TQCno_dRegEx.Test(FileName)
    'Debug.Print TQCdRegEx.Test(FileName)
    Is_RQC = Is_RQC Or RQCRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_LTR
''
'' Description:
''
'' Check is the input string (sample name) is a LTR.
''
'' Parameters:
''
''    FileName - Input string to check if it is a LTR
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "LTR", "Ltr", "ltr"
''
'' Examples:
''
'' --- Code
''    Dim LTRTestArray As Variant
''    'Dim LTRBKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    LTRTestArray = Array("LTR01.d", "LTR_01.d", "Ltr.d", "ltr.d", _
''                         "018_LTR-GroupA-01")
''
''    For arrayIndex = 0 To UBound(LTRTestArray) - LBound(LTRTestArray)
''        Debug.Print Sample_Type_Identifier.Is_LTR(CStr(LTRTestArray(arrayIndex))) & ": " & _
''                    LTRTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_LTR(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim LTRRegEx As RegExp
    Set LTRRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    LTRRegEx.Pattern = "(LTR|[Ll]tr)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_LTR = LTRRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_NIST
''
'' Description:
''
'' Check is the input string (sample name) is a NIST.
''
'' Parameters:
''
''    FileName - Input string to check if it is a NIST
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "NIST", "Nist", "nist"
''
'' Examples:
''
'' --- Code
''    Dim NISTTestArray As Variant
''    Dim arrayIndex As Integer
''
''    NISTTestArray = Array("NIST01.d", "NIST_01.d", "Nist.d", "nist.d", _
''                          "018_NIST-GroupA-01")
''
''    For arrayIndex = 0 To UBound(NISTTestArray) - LBound(NISTTestArray)
''        Debug.Print Sample_Type_Identifier.Is_NIST(CStr(NISTTestArray(arrayIndex))) & ": " & _
''                    NISTTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_NIST(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim NISTRegEx As RegExp
    Set NISTRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    NISTRegEx.Pattern = "(NIST|[Nn]ist)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_NIST = NISTRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_SRM
''
'' Description:
''
'' Check is the input string (sample name) is a SRM.
''
'' Parameters:
''
''    FileName - Input string to check if it is a SRM
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "SRM", "Srm", "srm"
''
'' Examples:
''
'' --- Code
''    Dim SRMTestArray As Variant
''    Dim arrayIndex As Integer
''
''    SRMTestArray = Array("SRM01.d", "SRM_01.d", "Srm.d", "srm.d", _
''                          "018_SRM-GroupA-01")
''
''    For arrayIndex = 0 To UBound(SRMTestArray) - LBound(SRMTestArray)
''        Debug.Print Sample_Type_Identifier.Is_SRM(CStr(SRMTestArray(arrayIndex))) & ": " & _
''                    SRMTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_SRM(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim SRMRegEx As RegExp
    Set SRMRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    SRMRegEx.Pattern = "(SRM|[Ss]rm)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_SRM = SRMRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_PBLK
''
'' Description:
''
'' Check is the input string (sample name) is a PBLK.
''
'' Parameters:
''
''    FileName - Input string to check if it is a PBLK
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "PBLK", "Pblk",
''    both "Extracted", "ISTD" and "Blanks" are presence,
''    both "Processed" and "Blanks" are presence.
''
'' Examples:
''
'' --- Code
''    Dim PBLKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    PBLKTestArray = Array("Blk_EXIS", _
''                          "006_Extracted Blank+ISTD01", "Istd-Extract BLK", _
''                          "ProcessBlank", "BlkProcessed_01", "001-PBLK")
''
''    For arrayIndex = 0 To UBound(PBLKTestArray) - LBound(PBLKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_PBLK(CStr(PBLKTestArray(arrayIndex))) & ": " & _
''                    PBLKTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_PBLK(ByVal FileName As String) As Boolean
    Dim BlankExISTDRegEx As RegExp
    Set BlankExISTDRegEx = New RegExp
    Dim ISTDBlankExRegEx As RegExp
    Set ISTDBlankExRegEx = New RegExp
    Dim ExBlankISTDRegEx As RegExp
    Set ExBlankISTDRegEx = New RegExp
    Dim ISTDExBlankRegEx As RegExp
    Set ISTDExBlankRegEx = New RegExp
    
    Dim PBlankRegEx As RegExp
    Set PBlankRegEx = New RegExp
    Dim ProcessedBlankRegEx As RegExp
    Set ProcessedBlankRegEx = New RegExp
    Dim BlankProcessedRegEx As RegExp
    Set BlankProcessedRegEx = New RegExp
 
    Dim BlankPattern As String
    Dim ExtractPattern As String
    Dim BlankExPattern As String
    Dim ExBlankPattern As String
    Dim ISTDPattern As String
    
    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(BL(AN)?K|[B,b]l(an)?k)"
    
    ExtractPattern = "EX(T|TR(ACT(ED)?)?)?|[Ee]x(t|tr(act(ed)?)?)?"
    
    'ISTD is compulsory if the word Processed is not used
    ISTDPattern = "(IS(TD)?|[Ii]std)"
    
    'Extract pattern though optional must be strictly around Blank pattern
    BlankExPattern = BlankPattern & "([\s,_,-]+)?" & ExtractPattern
    ExBlankPattern = ExtractPattern & "([\s,_,-]+)?" & BlankPattern
    
    BlankExISTDRegEx.Pattern = BlankExPattern & "([\s,_,-]+)?" & ISTDPattern
    ISTDBlankExRegEx.Pattern = ISTDPattern & "([\s,_,-]+)?" & BlankExPattern
    ExBlankISTDRegEx.Pattern = ExBlankPattern & "([\s,_,-,+]+)?" & ISTDPattern
    ISTDExBlankRegEx.Pattern = ISTDPattern & "([\s,_,-]+)?" & ExBlankPattern
    
    'We accept PBLK, strictly no space
    PBlankRegEx.Pattern = "P" & BlankPattern
    'We accept Processed_Blank
    ProcessedBlankRegEx.Pattern = "(PROCESS(ED)?)|([Pp]rocess(ed)?)" & "([\s,_,-]+)?" & BlankPattern
    'We do not accept BLK_P for the moment
    BlankProcessedRegEx.Pattern = BlankPattern & "([\s,_,-]+)?" & "(PROCESS(ED)?)|[Pp](rocess(ed)?)"
    
    If BlankExISTDRegEx.Test(FileName) Then
        Is_PBLK = True
    ElseIf ISTDBlankExRegEx.Test(FileName) Then
        Is_PBLK = True
    ElseIf ExBlankISTDRegEx.Test(FileName) Then
        Is_PBLK = True
    ElseIf ISTDExBlankRegEx.Test(FileName) Then
        Is_PBLK = True
    ElseIf PBlankRegEx.Test(FileName) Then
        Is_PBLK = True
    ElseIf ProcessedBlankRegEx.Test(FileName) Then
        Is_PBLK = True
    ElseIf BlankProcessedRegEx.Test(FileName) Then
        Is_PBLK = True
    Else
        Is_PBLK = False
    End If
    
End Function

'' Function: Is_BLK
''
'' Description:
''
'' Check is the input string (sample name) has the word "Blank"
'' Do note that it is considered as a UBLK only if it is not identified
'' as other blanks like PBLK, SBLK, etc...
''
'' Parameters:
''
''    FileName - Input string to check if it has the word "Blank"
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "Blank"
''
'' Examples:
''
'' --- Code
''    Dim BlankTestArray As Variant
''    Dim arrayIndex As Integer
''
''    BlankTestArray = Array("MyBlank", "BLK01.d", "Blk01.d", "51_blank_01.d", "07_BLANK-1.d", "001_BLK.d", "Blank_01.d", _
''                           "06122016_Liver_TAG_Even01_YongLiang_blk6.d", _
''                           "07062017_Marie_PL_MCF_blk_2.d", "Blk_PC-PS-SM_01.d", "Blank_1_08112016.d", _
''                           "blank_7-r001.d", "Blk_12.d", "041_blank.d", "20161006-pos-SIM-Blank-001.d", _
''                           "20170717-Coral lipids-dMRM-TAG-Blk-01.d", "20170623-pos-RP-blank-001.d", _
''                           "20170210-RPLC-Pos-Blank02.d", "131_BLK.d")
''
''    For arrayIndex = 0 To UBound(BlankTestArray) - LBound(BlankTestArray)
''        Debug.Print (Sample_Type_Identifier.Is_UBLK(CStr(BlankTestArray(arrayIndex)))) & ": " & _
''                    BlankTestArray(arrayIndex)
''        Debug.Print (Not (Sample_Type_Identifier.Is_PBLK(CStr(BlankTestArray(arrayIndex))))) & ": " & _
''                    BlankTestArray(arrayIndex)
''        Debug.Print (Not (Sample_Type_Identifier.Is_SBLK(CStr(BlankTestArray(arrayIndex))))) & ": " & _
''                    BlankTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_BLK(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim BlankRegEx As RegExp
    Set BlankRegEx = New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    'We need a whole word match as we do not want 'blank'ets
    BlankRegEx.Pattern = "(BL(AN)?K|[B,b]l(an)?k)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    'Debug.Print File
    Is_BLK = BlankRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_SBLK
''
'' Description:
''
'' Check is the input string (sample name) is a SBLK.
''
'' Parameters:
''
''    FileName - Input string to check if it is a SBLK
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "SBLK", "Sblk",
''    both "Solvent" and "Blanks" are presence
''
'' Examples:
''
'' --- Code
''    Dim SBLKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    SBLKTestArray = Array("SBLK", "Solvent_Blank", "SOL_BLK_001", _
''                          "006_solvent blank", "Solvent")
''
''    For arrayIndex = 0 To UBound(SBLKTestArray) - LBound(SBLKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_SBLK(CStr(SBLKTestArray(arrayIndex))) & ": " & _
''                    SBLKTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_SBLK(ByVal FileName As String) As Boolean
    Dim SolventBLKRegEx As RegExp
    Set SolventBLKRegEx = New RegExp
    Dim SBLKRegEx As RegExp
    Set SBLKRegEx = New RegExp
    
    Dim BlankPattern As String

    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(BL(AN)?K|[B,b]l(an)?k)"

    SBLKRegEx.Pattern = "S" & BlankPattern
    'We make the blank pattern optional if the word Solvent is present
    SolventBLKRegEx.Pattern = "SOL(VENT)?|[S,s]ol(vent)?" & "([\s,_,-]+)?" & BlankPattern & "?"

    If SBLKRegEx.Test(FileName) Then
        Is_SBLK = True
    ElseIf SolventBLKRegEx.Test(FileName) Then
        Is_SBLK = True
    Else
        Is_SBLK = False
    End If

End Function

'' Function: Is_MBLK
''
'' Description:
''
'' Check is the input string (sample name) is a MBLK.
''
'' Parameters:
''
''    FileName - Input string to check if it is a MBLK
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "MBLK", "Mblk",
''    both "Matrix" and "Blanks" are presence
''
'' Examples:
''
'' --- Code
''    Dim MBLKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    MBLKTestArray = Array("MBLK", "Matrix_Blank", "MATRIX_BLK_001", _
''                          "006_matrix blank", "Matrix")
''
''    For arrayIndex = 0 To UBound(MBLKTestArray) - LBound(MBLKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_MBLK(CStr(MBLKTestArray(arrayIndex))) & ": " & _
''                    MBLKTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_MBLK(ByVal FileName As String) As Boolean
    Dim MatrixBLKRegEx As RegExp
    Set MatrixBLKRegEx = New RegExp
    Dim MBLKRegEx As RegExp
    Set MBLKRegEx = New RegExp

    Dim BlankPattern As String

    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(BL(AN)?K|[B,b]l(an)?k)"

    MBLKRegEx.Pattern = "M" & BlankPattern
    'We make the blank pattern optional if the word Matrix is present
    MatrixBLKRegEx.Pattern = "MATRIX|[M,m]atrix" & "([\s,_,-]+)?" & BlankPattern & "?"

    If MBLKRegEx.Test(FileName) Then
        Is_MBLK = True
    ElseIf MatrixBLKRegEx.Test(FileName) Then
        Is_MBLK = True
    Else
        Is_MBLK = False
    End If

End Function

'' Function: Is_STD
''
'' Description:
''
'' Check is the input string (sample name) is a STD.
''
'' Parameters:
''
''    FileName - Input string to check if it is a STD
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "STD", "Std",
''    but not "ISTD"
''
'' Examples:
''
'' --- Code
''    Dim STDTestArray As Variant
''    Dim Not_STDTestArray As Variant
''    Dim arrayIndex As Integer
''
''    STDTestArray = Array("STD01.d", "STD_01.d", "Std.d", "std.d", _
''                          "018_STD-GroupA-01")
''
''    Not_STDTestArray = Array("ISTD01.d", "ISTD_01.d", "Istd.d", "istd.d", _
''                             "018_ISTD-GroupA-01")
''
''    For arrayIndex = 0 To UBound(STDTestArray) - LBound(STDTestArray)
''        Debug.Print Sample_Type_Identifier.Is_STD(CStr(STDTestArray(arrayIndex))) & ": " & _
''                    STDTestArray(arrayIndex)
''    Next
''
''    For arrayIndex = 0 To UBound(Not_STDTestArray) - LBound(Not_STDTestArray)
''        Debug.Print Sample_Type_Identifier.Is_STD(CStr(Not_STDTestArray(arrayIndex))) & ": " & _
''                    Not_STDTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_STD(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim STDRegEx As RegExp
    Set STDRegEx = New RegExp
    Dim ISTDRegEx As RegExp
    Set ISTDRegEx = New RegExp
    
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    STDRegEx.Pattern = "(STD|[Ss]td)"
    ISTDRegEx.Pattern = "(ISTD|[Ii]std)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_STD = STDRegEx.Test(OnlyLettersText) And Not (ISTDRegEx.Test(OnlyLettersText))
    
End Function

'' Function: Is_LQQ
''
'' Description:
''
'' Check is the input string (sample name) is a LQQ.
''
'' Parameters:
''
''    FileName - Input string to check if it is a LQQ
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "LQQ", "Lqq", "lqq"
''
'' Examples:
''
'' --- Code
''    Dim LQQTestArray As Variant
''    Dim arrayIndex As Integer
''
''    LQQTestArray = Array("LQQ01.d", "LQQ_01.d", "Lqq.d", "lqq.d", _
''                          "018_LQQ-GroupA-01")
''
''    For arrayIndex = 0 To UBound(LQQTestArray) - LBound(LQQTestArray)
''        Debug.Print Sample_Type_Identifier.Is_LQQ(CStr(LQQTestArray(arrayIndex))) & ": " & _
''                    LQQTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_LQQ(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim LQQRegEx As RegExp
    Set LQQRegEx = New RegExp
    
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    LQQRegEx.Pattern = "(LQQ|[Ll]qq)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_LQQ = LQQRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_CTRL
''
'' Description:
''
'' Check is the input string (sample name) is a CTRL.
''
'' Parameters:
''
''    FileName - Input string to check if it is a CTRL
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "CTRL", "Ctrl", "ctrl"
''
'' Examples:
''
'' --- Code
''    Dim CTRLTestArray As Variant
''    Dim arrayIndex As Integer
''
''    CTRLTestArray = Array("CTRL01.d", "CTRL_01.d", "Ctrl.d", "ctrl.d", _
''                          "018_CTRL-GroupA-01")
''
''    For arrayIndex = 0 To UBound(CTRLTestArray) - LBound(CTRLTestArray)
''        Debug.Print Sample_Type_Identifier.Is_CTRL(CStr(CTRLTestArray(arrayIndex))) & ": " & _
''                    CTRLTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_CTRL(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim CTRLRegEx As RegExp
    Set CTRLRegEx = New RegExp

    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    CTRLRegEx.Pattern = "(CTRL|[Cc]trl)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_CTRL = CTRLRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_DUP
''
'' Description:
''
'' Check is the input string (sample name) is a DUP.
''
'' Parameters:
''
''    FileName - Input string to check if it is a DUP
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "DUP", "Dup", "dup"
''
'' Examples:
''
'' --- Code
''    Dim DUPTestArray As Variant
''    Dim arrayIndex As Integer
''
''    DUPTestArray = Array("DUP01.d", "DUP_01.d", "Dup.d", "dup.d", _
''                          "018_DUP-GroupA-01")
''
''    For arrayIndex = 0 To UBound(DUPTestArray) - LBound(DUPTestArray)
''        Debug.Print Sample_Type_Identifier.Is_DUP(CStr(DUPTestArray(arrayIndex))) & ": " & _
''                    DUPTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_DUP(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim DUPRegEx As RegExp
    Set DUPRegEx = New RegExp

    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    DUPRegEx.Pattern = "(DUP|[Dd]up)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_DUP = DUPRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_SPIK
''
'' Description:
''
'' Check is the input string (sample name) is a SPIK.
''
'' Parameters:
''
''    FileName - Input string to check if it is a SPIK
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "SPIK", "Spik", "spik"
''
'' Examples:
''
'' --- Code
''    Dim SPIKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    SPIKTestArray = Array("SPIK01.d", "SPIK_01.d", "Spik.d", _
''                          "spik.d", "018_SPIK-GroupA-01", _
''                          "SPIKE01.d", "SPIKE_01.d", "Spike.d", _
''                          "spike.d", "018_SPIKE-GroupA-01")
''
''    For arrayIndex = 0 To UBound(SPIKTestArray) - LBound(SPIKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_SPIK(CStr(SPIKTestArray(arrayIndex))) & ": " & _
''                    SPIKTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_SPIK(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim SPIKRegEx As RegExp
    Set SPIKRegEx = New RegExp

    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    SPIKRegEx.Pattern = "(SPIK|[Ss]pik)"
    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_SPIK = SPIKRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_LTRBK
''
'' Description:
''
'' Check is the input string (sample name) is a LTRBK.
''
'' Parameters:
''
''    FileName - Input string to check if it is a LTRBK
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "LTRBK",
''    "LTR" and "No ISTD"
''
'' Examples:
''
'' --- Code
''    Dim LTRBKTestArray As Variant
''    Dim Not_LTRBKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    LTRBKTestArray = Array("LTRBK01.d", "LTRBK_01.d", "Ltrbk.d", _
''                           "ltrbk.d", "018_LTRBK-GroupA-01", _
''                           "LTRBLK01.d", "LTRBLK_01.d", "Ltrblk.d", _
''                           "ltrblk.d", "018_LTRBLK-GroupA-01", _
''                           "LTR_BLK01.d", "LTR_BLK_01.d", "Ltr_blk.d", _
''                           "ltr_blk.d", "018_LTR_BLK-GroupA-01", _
''                           "LTR-BLK01.d", "LTR-BLK_01.d", "Ltr-blk.d", _
''                           "ltr-blk.d", "018_LTR-BLK-GroupA-01", _
''                           "LTR BLK01.d", "LTR BLK_01.d", "Ltr blk.d", _
''                           "ltr blk.d", "018_LTR BLK-GroupA-01", _
''                           "No_ISTD_LTR.d", "NoIS_LTR.d", _
''                           "LTR_NoIS.d", "LTR-no_istd.d")
''    Not_LTRBKTestArray = Array("LTR01.d", "LTR_01.d", "Ltr.d", "ltr.d", _
''                               "018_LTR-GroupA-01", "NO_LTR.d")
''
''    For arrayIndex = 0 To UBound(LTRBKTestArray) - LBound(LTRBKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_LTRBK(CStr(LTRBKTestArray(arrayIndex))) & ": " & _
''                    LTRBKTestArray(arrayIndex)
''    Next
''
''    For arrayIndex = 0 To UBound(Not_LTRBKTestArray) - LBound(Not_LTRBKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_LTRBK(CStr(Not_LTRBKTestArray(arrayIndex))) & ": " & _
''                    Not_LTRBKTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_LTRBK(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim LTRBKRegEx As RegExp
    Set LTRBKRegEx = New RegExp
    Dim NoISTDLTRRegEx As RegExp
    Set NoISTDLTRRegEx = New RegExp
    Dim LTRNoISTDRegEx As RegExp
    Set LTRNoISTDRegEx = New RegExp

    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    Dim BlankPattern As String
    Dim NoISTDPattern As String

    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(B(L|LAN)?K|[B,b](l|lan)?k)"
    
    'No ISTD pattern
    NoISTDPattern = "(NO|[Nn]o)" & "([\s,_,-]+)?" & "(IS(TD)?|[Ii]std)"
    
    'We make the blank pattern
    LTRBKRegEx.Pattern = "(LTR|[Ll]tr)" & "([\s,_,-]+)?" & BlankPattern
    NoISTDLTRRegEx.Pattern = NoISTDPattern & "([\s,_,-]+)?" & "(LTR|[Ll]tr)"
    LTRNoISTDRegEx.Pattern = "(LTR|[Ll]tr)" & "([\s,_,-]+)?" & NoISTDPattern

    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_LTRBK = LTRBKRegEx.Test(OnlyLettersText) Or _
              NoISTDLTRRegEx.Test(OnlyLettersText) Or _
              LTRNoISTDRegEx.Test(OnlyLettersText)
    
End Function

'' Function: Is_NISTBK
''
'' Description:
''
'' Check is the input string (sample name) is a NISTBK.
''
'' Parameters:
''
''    FileName - Input string to check if it is a NISTBK
''
'' Returns:
''    A boolean (True or False). Return True if
''    the input string contains "NISTBK",
''    "NIST" and "No ISTD"
''
'' Examples:
''
'' --- Code
''    Dim NISTBKTestArray As Variant
''    Dim Not_NISTBKTestArray As Variant
''    Dim arrayIndex As Integer
''
''    NISTBKTestArray = Array("NISTBK01.d", "NISTBK_01.d", "Nistbk.d", _
''                            "nistbk.d", "018_NISTBK-GroupA-01", _
''                            "NISTBLK01.d", "NISTBLK_01.d", "Nistblk.d", _
''                            "nistblk.d", "018_NISTBLK-GroupA-01", _
''                            "NIST_BLK01.d", "NIST_BLK_01.d", "Nist_blk.d", _
''                            "nist_blk.d", "018_NIST_BLK-GroupA-01", _
''                            "NIST-BLK01.d", "NIST-BLK_01.d", "Nist-blk.d", _
''                            "nist-blk.d", "018_NIST-BLK-GroupA-01", _
''                            "NIST BLK01.d", "NIST BLK_01.d", "Nist blk.d", _
''                            "nist blk.d", "018_NIST BLK-GroupA-01", _
''                            "No_ISTD_NIST.d", "NoIS_NIST.d", _
''                            "NIST_NoIS.d", "NIST-no_istd.d")
''    Not_NISTBKTestArray = Array("NIST01.d", "NIST_01.d", "Nist.d", "nist.d", _
''                                "018_NIST-GroupA-01", "NO_NIST.d")
''
''    For arrayIndex = 0 To UBound(NISTBKTestArray) - LBound(NISTBKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_NISTBK(CStr(NISTBKTestArray(arrayIndex))) & ": " & _
''                    NISTBKTestArray(arrayIndex)
''    Next
''
''    For arrayIndex = 0 To UBound(Not_NISTBKTestArray) - LBound(Not_NISTBKTestArray)
''        Debug.Print Sample_Type_Identifier.Is_NISTBK(CStr(Not_NISTBKTestArray(arrayIndex))) & ": " & _
''                    Not_NISTBKTestArray(arrayIndex)
''    Next
'' ---
Public Function Is_NISTBK(ByVal FileName As String) As Boolean
    Dim NonLettersRegEx As RegExp
    Set NonLettersRegEx = New RegExp
    Dim NISTBKRegEx As RegExp
    Set NISTBKRegEx = New RegExp
    Dim NoISTDNISTRegEx As RegExp
    Set NoISTDNISTRegEx = New RegExp
    Dim NISTNoISTDRegEx As RegExp
    Set NISTNoISTDRegEx = New RegExp

    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    Dim BlankPattern As String
    Dim NoISTDPattern As String

    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(B(L|LAN)?K|[B,b](l|lan)?k)"
    
    'No ISTD pattern
    NoISTDPattern = "(NO|[Nn]o)" & "([\s,_,-]+)?" & "(IS(TD)?|[Ii]std)"
    
    'We make the blank pattern
    NISTBKRegEx.Pattern = "(NIST|[Nn]ist)" & "([\s,_,-]+)?" & BlankPattern
    NoISTDNISTRegEx.Pattern = NoISTDPattern & "([\s,_,-]+)?" & "(NIST|[Nn]ist)"
    NISTNoISTDRegEx.Pattern = "(NIST|[Nn]ist)" & "([\s,_,-]+)?" & NoISTDPattern

    OnlyLettersText = Trim$(NonLettersRegEx.Replace(FileName, " "))
    Is_NISTBK = NISTBKRegEx.Test(OnlyLettersText) Or _
              NoISTDNISTRegEx.Test(OnlyLettersText) Or _
              NISTNoISTDRegEx.Test(OnlyLettersText)
    
End Function
