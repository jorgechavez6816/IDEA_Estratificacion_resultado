Sub Main
	Call Stratification()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Análisis: Estratificación
Function Stratification
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Stratification
	task.IncludeAllFields
	resultName = db.UniqueResultName("Estratificacion")
	task.ResultName = resultName
	task.FieldToStratify = "TOTAL"
	task.AddFieldToTotal "SUMA_TOTAL"
	task.LowerLimit -52.71
	task.AddUpperLimit 9947.29
	task.AddUpperLimit 19947.29
	task.AddUpperLimit 29947.29
	task.AddUpperLimit 39947.29
	task.AddUpperLimit 49947.29
	task.AddUpperLimit 59947.29
	task.AddUpperLimit 69947.29
	task.AddUpperLimit 79947.29
	task.AddUpperLimit 89947.29
	task.AddUpperLimit 99947.29
	task.CreateVirtualDatabase = False
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Function