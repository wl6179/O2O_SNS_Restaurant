# O2O_SNS_Restaurant 社会化餐厅（ASP）
从技术上，这已经算是我较高水平的一个ASP技术作品了，也拥有比较完整的可重用性，有一定的模式，也已经能轻松完成较复杂的系统，所以可以做为我ASP研究的一个里程碑，而更高水平用在公司里不方便开源。2010 年底也是我开始探索并使用 PHP 技术的时间点，为逃离盗版，因为为餐厅推荐安装盗版软件是一件很不舒服的事情。

- 注意：index.Welcome 的后缀 *.Welcome 是为了用户体验，实则 ASP 程序
- 注意2：为了更好的保护客户的隐私信息，程序中的所有图片已经做好了清除处理，只有程序
- 拥有完整的会员体系
- 拥有完整的积分系统
- 拥有向会员发放优惠券功能，并支持客户自行打印上餐厅消费（时代限制）
- 拥有电子邮件通讯能力
- 会员生日提醒功能，餐厅可向客户送出生日优惠券
- 会员点评餐厅菜品
- 会员专区，可展示积分换购会员的动态，还能显示许多快过生日的会员榜单，送去祝福
- 餐厅类小游戏专区
- 会员之间浏览点赞
- 餐品展示+点评菜品星级，菜品推荐，时令菜推荐，新菜上线展示
- 几乎就是 SNS 的餐厅网站，需要庞大而复杂的 SQL设计工作 + 还很需要当时极时髦的 JSON 格式通讯来完成 Facebook 式的体验
- 几乎就是 O2O 的雏形，让会员享受到及时的优惠券，免费喝可乐、打8.5折都不在话下
- 别忘了，这个时代还在谈论凡客诚品，O2O 一词还尚未出现
- 此时，我已经关注扫条码机、收款机、及针对监控采集前厅电话线的开发了，以便创建出超级会员系统。但是此时还是诺基亚的天下，没有安卓系统体系，所有迹象表明只能重新开始研究 C、Delphi、Visual Basic，而终结！
- 我是具有探索精神的实践者、总设计师：Chris Wang
- 2010 年 ASP 作品
- 工具：CVSNT Server 古董级的版本质量控制、WinCVS、Windows Server、SQL Server 2005 最新窗口函数（SNS必用！）、Dreamweaver、ASP、Visual Basic、dojo 1.4.1 框架、Microsoft Web Application Stress Tool 微软性能测试工具 等

例子1 - 窗口函数：
````sql
--最受欢迎菜
select top 7 *,   --这个 * 在现在看来，是性能隐患~
  (
  select distinct cast(sumChineseDish_Taste as decimal)/TotalChineseDish_Taste as avgChineseDish_Taste 
  from (select product_id, 
          sum(ChineseDish_Taste) over() as sumChineseDish_Taste,    --这就是 SQL Server 2005 最新的新特性 - 窗口函数
          count(ChineseDish_Taste) over() as TotalChineseDish_Taste 
        from [CXBG_account_RemarkOn] 
        where product_id=[CXBG_product].id 
          and deleted=0 and ChineseDish_Taste>0
        ) as x
  ) as avgChineseDish_TasteNow 
from [CXBG_product] 
where 
  deleted=0 and isOnsale=1 
order by avgChineseDish_TasteNow desc,OrderID desc,id desc
````

例子2 - 窗口函数：
````sql
--首页新品推荐[套餐]
select top 3 *,		--得出星评率最高e前3名（的菜品）！
	(
	select distinct cast(sumStarRating as decimal)/TotalStarRating as avgStarRating		--计算所有会员对每一道菜品的星评率
	from (select product_id,
			  sum(theStarRatingForChineseDishInformation) over() as sumStarRating,	--统计本菜的评星总数
			  count(theStarRatingForChineseDishInformation)over() as TotalStarRating	--统计本菜的总评人数
		  from [CXBG_account_RemarkOn] 
		  where product_id=[CXBG_product].id 
			  and deleted=0	--有效的食评
			  and theStarRatingForChineseDishInformation>0	--大于等于1星的（有效的）食评
		  ) as x
	) as avgStarRatingNow 
from [CXBG_product] 
where 
  deleted=0 and isOnsale=1 and isSetMeals=1 and 1=1	--是套餐
order by avgStarRatingNow desc,OrderID desc,id desc	-- * （关键）以"星评率"作排序！
````
