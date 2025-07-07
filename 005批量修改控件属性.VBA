Sub ChangeTrigger()
On Error Resume Next

Dim objResult, objMEBER, objCount

Set objResult = ActkveDlcjment.HMIObjects.Find(objectName:="*",objectTpye:="HMICircle")
Set ojbCount = objResult.Count

Msgbox "对象的数量" & CStr(ojbcount)

For Each objMEBER In objResult

objMEBER.Properties("BackColor").Dynamc.trigger.VariableTrigger(1).CycleType = 0
objMEBER.Selected = True

Next objMEBER

Msgbox "操作完成"

End Sub

