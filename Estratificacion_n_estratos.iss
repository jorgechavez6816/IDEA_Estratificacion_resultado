Sub Main
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set stats = db.FieldStats("TOTAL") 		'Obtener las estadísticas de campo.
	n1 = 10					' Obtención de estratos
	a1 = (stats.MaxValue())
	a2 = (stats.MinValue())
	a3 = (Abs(a1)+Abs(a2))/n1
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.Stratification
	resultName = db.UniqueResultName("Estratificación100")
	task.ResultName = resultName
	task.FieldToStratify = "TOTAL"
	task.AddFieldToTotal "TOTAL"
	task.LowerLimit (a2)
	task.AddUpperLimit (a3*1)
	task.AddUpperLimit (a3*2)
	task.AddUpperLimit (a3*3)
	task.AddUpperLimit (a3*4)
	task.AddUpperLimit (a3*5)
	task.AddUpperLimit (a3*6)
	task.AddUpperLimit (a3*7)
	task.AddUpperLimit (a3*8)
	task.AddUpperLimit (a3*9)
	task.AddUpperLimit (a3*10)
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
End Sub



