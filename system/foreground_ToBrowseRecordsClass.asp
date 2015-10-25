<%
'模块说明：前台专用的菜品浏览历史记录类库.
'日期说明：2010-1-24
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀技术有限公司。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>
<%
Class ToBrowseRecordsClass
	'*****************************************成员变量*****************************************
	'Private intNum
	'*****************************************************************************************
	
	'*****************************************GET 属性*****************************************
	'获取菜品浏览历史的数目.(字典中的项目数).只读属性.
	Public Property Get Count()
		Count=Session("ToBrowseRecords").Count
	End Property
	'*****************************************************************************************
	
	'*****************************************LET 属性*****************************************
	'修改...属性.
'	Public Property LET Num(i)
'		intNum=i
'	End Property
	'*****************************************************************************************
	
	'初始化.
	Private Sub Class_Initialize()
		Call CreateToBrowseRecords()
		'记住不要在初始化中放入太多的过程！！！
    End Sub
	
	'结束化.
	Private Sub Class_Terminate()
    	'不要马上删除字典对象，因为整个大会话过程都还需要用得着访问它。
		'Set Session("ToBrowseRecords")=Nothing
    End Sub
	
	
	'*****************************************方法********************************************
	'创建菜品浏览历史.
	Public Sub CreateToBrowseRecords()
		If IsObject(Session("ToBrowseRecords")) Then
			'Response.Write "<br />Session(""ToBrowseRecords"")已经是对象！"
			
		Else
			Set Session("ToBrowseRecords")=Server.CreateObject("Scripting.Dictionary")
			'Response.Write "<br />重新创建Session(""ToBrowseRecords"")新对象！"
		End If
	End Sub
	
	'清除菜品浏览历史.
	Public Sub ClearToBrowseRecords()
		If IsObject(Session("ToBrowseRecords")) Then
			Session("ToBrowseRecords").RemoveAll
			'Response.Write "<br />Session(""ToBrowseRecords"")的项目清除为空！"
		End If
		
	End Sub
	
	'完全销毁对象.
	Public Sub CloseToBrowseRecords()
		If IsObject(Session("ToBrowseRecords")) Then
			Set Session("ToBrowseRecords")	=Nothing
			Session("ToBrowseRecords")		=Null
			'Response.Write "<br />完全销毁Session(""ToBrowseRecords"")对象！"
		End If
		
	End Sub
	
	
	'获取菜品的ID的集合，用于传输参数，以及记入Cookies.
	Public Function IDs()
		'获取菜品的ID的集合.
		Dim key_1
		For Each key_1 In Session("ToBrowseRecords").Keys
			IDs=IDs & key_1 &","
		Next
		If Instr(IDs,",")>0 Then IDs=Left(IDs, Len(IDs)-1)
	End Function
	
	'检测菜品浏览历史对象是否存在.
	Public Function CheckToBrowseRecords()
		If IsObject(Session("ToBrowseRecords")) Then
			CheckToBrowseRecords=True
		Else
			CheckToBrowseRecords=False
		End If
	End Function
	
	'检测某一个菜品浏览历史（菜品）.
	Public Function CheckProduct(pID)
		CheckProduct=False
		If isNumeric(pID) Then
			pID=CokeShow.CokeClng(pID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class ToBrowseRecordsClass", "方法CheckProduct的参数pID不正确，无法检测某一件菜品的操作！"
			Exit Function
		End If
		
		If CheckToBrowseRecords=True Then
			If Session("ToBrowseRecords").Exists(pID)=True Then
				CheckProduct=True
			End If
		End If
	End Function
	
	'移除某一菜品浏览历史.
	Public Function RemoveProduct(pID)
		RemoveProduct=False
		If isNumeric(pID) Then
			pID=CokeShow.CokeClng(pID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class ToBrowseRecordsClass", "方法RemoveProduct的参数pID不正确，无法移除某一件菜品的操作！"
			Exit Function
		End If
		If Not CheckProduct(pID) Then
			'此菜品不存在，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class ToBrowseRecordsClass", "方法RemoveProduct的CheckProduct(pID)检测当前pID无对应项目，无法移除某一件菜品的操作！"
			Exit Function
		End If
		
		If Session("ToBrowseRecords").Exists(pID)=True Then
			Session("ToBrowseRecords").Remove(pID)
			RemoveProduct=True
		End If
	End Function
	
	'新增一件菜品.
	'注意：参数strName_strUnitPrice：必须是这样的格式 ProductName$$UnitPrice$$Photo 即整个调用过程为： 
	'= ToBrowseRecords.AddProduct(8, "微软键盘套装MS-DSN$$125.00$$/images/NoPic.gif").
	Public Function AddProduct(pID,strName_strUnitPrice)
		AddProduct=False
		If isNumeric(pID) Then
			pID=CokeShow.CokeClng(pID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class ToBrowseRecordsClass", "方法AddProduct的参数pID不正确，无法进行字典Add操作！"
			Exit Function
		End If
		If Len(strName_strUnitPrice)>0 And Instr(strName_strUnitPrice,"$$")>0 Then
			strName_strUnitPrice=Trim(strName_strUnitPrice)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class ToBrowseRecordsClass", "方法AddProduct的参数strName_strUnitPrice不正确，无法新增一件菜品的操作！"
			Exit Function
		End If
		
		If CheckProduct(pID)=True Then
			'如果已经有此菜品，那么不做操作.更新一下值.
			Session("ToBrowseRecords").Item(pID) = strName_strUnitPrice
		Else
			'新增此菜品.
			'使用类的时候，已经想办法自动创建字典对象了.
			Session("ToBrowseRecords").Add  pID,strName_strUnitPrice
		End If
		AddProduct=True
		
	End Function
	
	'调用获取某一菜品的某个属性值.
	'调用方法：=ToBrowseRecords.GetProductValue(3,"ProductName") 或者 UnitPrice-----表示3号菜品的Name字段或者UnitPrice字段的值！
	Public Function GetProductValue(pID,strSomeField)
		GetProductValue=False
		If isNumeric(pID) Then
			pID=CokeShow.CokeClng(pID)
		Else
			'参数不正确，退出.操作失败.
			'Response.Clear()
			Err.Raise vbObjectError + 1288, "Class ToBrowseRecordsClass", "方法GetProductValue的参数pID不正确，无法获取某菜品其它字段值的操作！"
			Exit Function
		End If
		
		'拆分字段，取出需要的值.
		If CheckProduct(pID)=True Then
			Select Case strSomeField
				Case "ProductName"
					GetProductValue=Split( Session("ToBrowseRecords").Item(pID), "$$" )(0)		'项目的值.
				Case "UnitPrice"
					GetProductValue=Split( Session("ToBrowseRecords").Item(pID), "$$" )(1)		'项目的值.
				Case "Photo"
					GetProductValue=Split( Session("ToBrowseRecords").Item(pID), "$$" )(2)		'项目的值.
			End Select
		Else
			'取不到结果返回False.
			GetProductValue=False
		End If
		
	End Function
	'*****************************************************************************************
End Class

%>