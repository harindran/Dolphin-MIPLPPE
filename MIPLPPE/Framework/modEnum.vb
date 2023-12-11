Module modEnum
    Public definenew As Boolean = False
    Public FGValueFNC, FGValueINR As Double
    Public FGDocEntry, RDate, DEFINE As String
    Public RMValueFNC, RMValueINR, StdFNCValue, StdINRValue, ReceiptFNC, FBDINR As Double
    Public BDFNC, BDINR As Double
    Public RMDocEntry As String
    Public RMLineId, FGLineId, FGLineNO, RMLineNO, ReceiptNo, BDocEntry, BLineId As Integer
    Public ct, tt As String
    Public Enum LinkedType
        LinkedSystemObject = 1
        LinkedTable = 2
        LinkedUDO = 3
    End Enum
End Module
