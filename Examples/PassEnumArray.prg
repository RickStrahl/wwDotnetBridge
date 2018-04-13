
do wwDotNetBridge
LOCAL loBridge as wwDotNetBridge
loBridge = CreateObject("wwDotNetBridge","V4")
? loBridge.LoadAssembly("System.Windows.Forms")

loTest = loBridge.CreateInstance("Westwind.WebConnection.TypePassingTests")
? loTest

LOCAL loArray as Westwind.WebConnection.ComArray
loArray = loBridge.Createarray("System.Windows.Forms.MessageBoxButtons")

loComValue = loBridge.CreateComValue()
loComValue.SetEnum("System.Windows.Forms.MessageBoxButtons.OK")
loArray.AddItem( loComValue) 

loComValue = loBridge.CreateComValue()
loComValue.SetEnum("System.Windows.Forms.MessageBoxButtons.OKCancel")
loArray.AddItem( loComValue) 

loComValue = loBridge.CreateComValue()
loComValue.SetEnum("System.Windows.Forms.MessageBoxButtons.YesNo")
loArray.AddItem( loComValue) 

? loArray.Count()
? loBridge.InvokeMethod(loTest,"PassEnumArray",loArray)
RETURN