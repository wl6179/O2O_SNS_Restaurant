<%
'模块说明：图形操作类库.
'日期说明：2009-11-26
'版权说明：www.cokeshow.com.cn	可乐秀(北京)技术有限公司产品
'友情提示：本产品已经申请专利，请勿用于商业用途，如果需要帮助请联系可乐秀技术有限公司。
'联系电话：(010)-67659219
'电子邮件：cokeshow@qq.com
%>
<%
Class PhotosClass
	Public Jpeg
	Private ABC
	
	
	Private Sub Class_Initialize()
		'If Not IsObject(Jpeg) Then Set Jpeg=Server.CreateObject("Persits.Jpeg")    '创建实例
    End Sub
	
	Private Sub Class_Terminate()
    	On Error Resume Next
			If IsObject(Jpeg) Then Set Jpeg=Nothing
    End Sub
	
	Public Sub Start()
        
    End Sub
	
	
	
	
'生成图片区
'*************************************************************************
	'生成图片缩略图函数.
	'参数：
	'1.thePhotoPath		图片读取路径
	'2.finalWidth		要缩小的尺寸（最终宽度），然后对进行等比例缩小
	'3.theGeneratePath	图片生成路径
	'例如: CokeShowPhotosClass.shortPhoto("/uploadimages/test1.jpg", 120, "/uploadimages/120/test1.jpg")
	Public Sub shortPhoto(thePhotoPath, finalWidth, theGeneratePath)
		Dim Scale
		'验证参数
		If isNull(thePhotoPath) Then Exit Sub
		If Not isNumeric(finalWidth) Then Exit Sub
		If isNull(theGeneratePath) Then Exit Sub
		
		'打开图片
		Set Jpeg=Server.CreateObject("Persits.Jpeg")
		Jpeg.Open Server.MapPath(thePhotoPath)		'处理图片路径，如"savephoto/test1.jpg".
		'计算出宽度变化的比例
		Scale = finalWidth / Jpeg.OriginalWidth
		'调整宽度和高度
		Jpeg.Width	= Jpeg.OriginalWidth * Scale
		Jpeg.Height	= Jpeg.OriginalHeight * Scale
		
		'设置压缩率
		Jpeg.Quality=90
		'设定锐化效果
		Jpeg.Sharpen 1, 120
		'Jpeg.Canvas.Font.Quality=4
		
		'水印相关
		'Jpeg.Canvas.Font.Color = &H000000 ''''//水印字体颜色
		'Jpeg.Canvas.Font.Family = "宋体" ''''//水印字体
		'Jpeg.Canvas.Font.Size = 14 ''''//水印字体大小
		'Jpeg.Canvas.Font.Bold = False ''''//是否粗体，粗体用：True
		'Jpeg.Canvas.Font.BkMode = &HFFFFFF ''''//字体背景颜色
		'Jpeg.Canvas.Print 10, 10, "www.chinateeyoo.com" ''''//水印文字，两个数字10为水印的xy座标
		
		'保存新图片.
		Jpeg.Save Server.MapPath(theGeneratePath)	'保存路径
		'销毁对象.
		Jpeg.Close
		
	End Sub
		
'*************************************************************************
End Class
%>