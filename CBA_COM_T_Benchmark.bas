Attribute VB_Name = "CBA_COM_T_Benchmark"
'Callback for CBA_COM_T_Benchmark onAction
Sub CBA_COM_T_BenchmarkRun(Control As IRibbonControl)
Dim F As fCOM_Bench
    Set F = New fCOM_Bench
    If F.bIsActive Then F.Show vbModeless
End Sub
