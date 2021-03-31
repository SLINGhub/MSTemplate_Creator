Attribute VB_Name = "Sample_Type_Identifier"
Public Function Get_Sample_Type(FileName As String) As String
    If isEQC(FileName) Then
        Get_Sample_Type = "EQC"
    ElseIf isSST(FileName) Then
        Get_Sample_Type = "SST"
    ElseIf isBQC(FileName) Then
        Get_Sample_Type = "BQC"
    ElseIf isTQC(FileName) And Not isRQC(FileName) Then
        Get_Sample_Type = "TQC"
    ElseIf isRQC(FileName) Then
        Get_Sample_Type = "RQC"
    ElseIf isLTR(FileName) And Not isLTRBK(FileName) Then
        Get_Sample_Type = "LTR"
    ElseIf isNIST(FileName) And Not isNISTBK(FileName) Then
        Get_Sample_Type = "NIST"
    ElseIf isSRM(FileName) Then
        Get_Sample_Type = "SRM"
    ElseIf isPBLK(FileName) Then
        Get_Sample_Type = "PBLK"
    ElseIf isUBLK(FileName) And Not isPBLK(FileName) And Not isSBLK(FileName) _
        And Not isMBLK(FileName) And Not isLTRBK(FileName) _
        And Not isNISTBK(FileName) Then
        Get_Sample_Type = "UBLK"
    ElseIf isSBLK(FileName) Then
        Get_Sample_Type = "SBLK"
    ElseIf isMBLK(FileName) Then
        Get_Sample_Type = "MBLK"
    ElseIf isSTD(FileName) Then
        Get_Sample_Type = "STD"
    ElseIf isLQQ(FileName) Then
        Get_Sample_Type = "LQQ"
    ElseIf isCTRL(FileName) Then
        Get_Sample_Type = "CTRL"
    ElseIf isDUP(FileName) Then
        Get_Sample_Type = "DUP"
    ElseIf isSPIK(FileName) Then
        Get_Sample_Type = "SPIK"
    ElseIf isLTRBK(FileName) Then
        Get_Sample_Type = "LTRBK"
    End If
    
End Function

Public Function isEQC(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim EQCRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    EQCRegEx.Pattern = "(EQC|[Ee]qc)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isEQC = EQCRegEx.Test(OnlyLettersText)
    
End Function

Public Function isSST(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim SSTRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    SSTRegEx.Pattern = "(SST|[Ss]st)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isSST = SSTRegEx.Test(OnlyLettersText)
End Function

Public Function isBQC(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim BQCRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    BQCRegEx.Pattern = "([BP]QC|[Pp]qc|[Bb]qc)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isBQC = BQCRegEx.Test(OnlyLettersText)
    
End Function

Public Function isTQC(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim TQCRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    TQCRegEx.Pattern = "(TQC|[Tt]qc)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isTQC = TQCRegEx.Test(OnlyLettersText)
    
End Function

Public Function isRQC(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim TQCdRegEx As New RegExp
    Dim TQCno_dRegEx As New RegExp
    Dim RQCRegEx As New RegExp
    Dim PercentRegEx As New RegExp
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
    
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    RQCRegEx.Pattern = "(RQC|[Rr]qc)"
    
    isRQC = TQCdRegEx.Test(FileName) Or TQCno_dRegEx.Test(FileName)
    'Debug.Print TQCdRegEx.Test(FileName)
    isRQC = isRQC Or RQCRegEx.Test(OnlyLettersText)
    
End Function

Public Function isLTR(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim LTRRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    LTRRegEx.Pattern = "(LTR|[Ll]tr)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isLTR = LTRRegEx.Test(OnlyLettersText)
    
End Function

Public Function isNIST(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim NISTRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    NISTRegEx.Pattern = "(NIST|[Nn]ist)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isNIST = NISTRegEx.Test(OnlyLettersText)
    
End Function

Public Function isSRM(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim SRMRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    SRMRegEx.Pattern = "(SRM|[Ss]rm)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isSRM = SRMRegEx.Test(OnlyLettersText)
    
End Function

Public Function isPBLK(FileName As String) As Boolean

    Dim BlankExISTDRegEx As New RegExp
    Dim ISTDBlankExRegEx As New RegExp
    Dim ExBlankISTDRegEx As New RegExp
    Dim ISTDExBlankRegEx As New RegExp
    
    Dim PBlankRegEx As New RegExp
    Dim ProcessedBlankRegEx As New RegExp
    Dim BlankProcessedRegEx As New RegExp
    
    Dim BlankPattern As String
    Dim ExtractPattern As String
    Dim BlankExPattern As String
    Dim ExBlankPattern As String
    Dim ISTDPattern As String
    Dim ProcessPattern As String
    
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
        isPBLK = True
    ElseIf ISTDBlankExRegEx.Test(FileName) Then
        isPBLK = True
    ElseIf ExBlankISTDRegEx.Test(FileName) Then
        isPBLK = True
    ElseIf ISTDExBlankRegEx.Test(FileName) Then
        isPBLK = True
    ElseIf PBlankRegEx.Test(FileName) Then
        isPBLK = True
    ElseIf ProcessedBlankRegEx.Test(FileName) Then
        isPBLK = True
    ElseIf BlankProcessedRegEx.Test(FileName) Then
        isPBLK = True
    Else
        isPBLK = False
    End If
    
End Function

Public Function isUBLK(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim BlankRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    'We need a whole word match as we do not want 'blank'ets
    BlankRegEx.Pattern = "(BL(AN)?K|[B,b]l(an)?k)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    'Debug.Print File
    isUBLK = BlankRegEx.Test(OnlyLettersText)
    
End Function

Public Function isSBLK(FileName As String) As Boolean
    Dim SolventBLKRegEx As New RegExp
    Dim SBLKRegEx As New RegExp
    Dim BlankPattern As String

    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(BL(AN)?K|[B,b]l(an)?k)"

    SBLKRegEx.Pattern = "S" & BlankPattern
    'We make the blank pattern optional if the word Solvent is present
    SolventBLKRegEx.Pattern = "SOL(VENT)?|[S,s]ol(vent)?" & "([\s,_,-]+)?" & BlankPattern & "?"

    If SBLKRegEx.Test(FileName) Then
        isSBLK = True
    ElseIf SolventBLKRegEx.Test(FileName) Then
        isSBLK = True
    Else
        isSBLK = False
    End If

End Function

Public Function isMBLK(FileName As String) As Boolean
    Dim MatrixBLKRegEx As New RegExp
    Dim MBLKRegEx As New RegExp
    Dim BlankPattern As String

    'We need a whole word match as we do not want 'blank'ets
    BlankPattern = "(BL(AN)?K|[B,b]l(an)?k)"

    MBLKRegEx.Pattern = "M" & BlankPattern
    'We make the blank pattern optional if the word Matrix is present
    MatrixBLKRegEx.Pattern = "MATRIX|[M,m]atrix" & "([\s,_,-]+)?" & BlankPattern & "?"

    If MBLKRegEx.Test(FileName) Then
        isMBLK = True
    ElseIf MatrixBLKRegEx.Test(FileName) Then
        isMBLK = True
    Else
        isMBLK = False
    End If

End Function

Public Function isSTD(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim STDRegEx As New RegExp
    Dim ISTDRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    STDRegEx.Pattern = "(STD|[Ss]td)"
    ISTDRegEx.Pattern = "(ISTD|[Ii]std)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isSTD = STDRegEx.Test(OnlyLettersText) And Not (ISTDRegEx.Test(OnlyLettersText))
    
End Function

Public Function isLQQ(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim LQQRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    LQQRegEx.Pattern = "(LQQ|[Ll]qq)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isLQQ = LQQRegEx.Test(OnlyLettersText)
    
End Function

Public Function isCTRL(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim CTRLRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    CTRLRegEx.Pattern = "(CTRL|[Cc]trl)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isCTRL = CTRLRegEx.Test(OnlyLettersText)
    
End Function

Public Function isDUP(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim DUPRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    DUPRegEx.Pattern = "(DUP|[Dd]up)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isDUP = DUPRegEx.Test(OnlyLettersText)
    
End Function

Public Function isSPIK(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim SPIKRegEx As New RegExp
    Dim OnlyLettersText As String
    NonLettersRegEx.Pattern = "[^A-Za-z]"
    NonLettersRegEx.Global = True
    
    SPIKRegEx.Pattern = "(SPIK|[Ss]pik)"
    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isSPIK = SPIKRegEx.Test(OnlyLettersText)
    
End Function

Public Function isLTRBK(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim LTRBKRegEx As New RegExp
    Dim NoISTDLTRRegEx As New RegExp
    Dim LTRNoISTDRegEx As New RegExp
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

    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isLTRBK = LTRBKRegEx.Test(OnlyLettersText) Or _
              NoISTDLTRRegEx.Test(OnlyLettersText) Or _
              LTRNoISTDRegEx.Test(OnlyLettersText)
    
End Function

Public Function isNISTBK(FileName As String) As Boolean
    Dim NonLettersRegEx As New RegExp
    Dim NISTBKRegEx As New RegExp
    Dim NoISTDNISTRegEx As New RegExp
    Dim NISTNoISTDRegEx As New RegExp
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

    OnlyLettersText = Trim(NonLettersRegEx.Replace(FileName, " "))
    isNISTBK = NISTBKRegEx.Test(OnlyLettersText) Or _
              NoISTDNISTRegEx.Test(OnlyLettersText) Or _
              NISTNoISTDRegEx.Test(OnlyLettersText)
    
End Function
