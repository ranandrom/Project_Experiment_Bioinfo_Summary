使用该程序前请仔细阅读：

1.该程序主要功能：为实现把指定目录下的实验数据文件（.xlsx格式文件）与指定目录下的生信数据文件（.xlsx格式文件）进行合并数据，然后生成新的合并文件（.xlsx格式文件）存放到指定目录下。

2.当前版本：V1.0.0。

3.该程序的运行环境是：linux系统+java1.8版本以上（目前规定为本公司的office_server服务器上运行）。

4.该程序的使用方法为：java -jar Project_Experiment_Bioinfo_Summary.jar xxxx(输入参数)
其中输入参数为：
(1).【-B/-b 指定的生信文件所在路径】：如果不输入指定时，默认为：/wdmycloud/anchordx_cloud/杨莹莹/项目-生信-汇总表/（最新日期命名的目录）。
(2).【-E/-e 指定的实验文件所在路径】：如果不输入指定时，默认为：/wdmycloud/anchordx_cloud/杨莹莹/项目进展汇总表。
(3).【-O/-o 指定的文件输出路径】：如果不输入指定时，默认为：/wdmycloud/anchordx_cloud/杨莹莹/项目实验生信汇总表/（以当天日期命名的目录）。

5.作者姓名：卢志荣
  邮箱：zhirong_lu@anchordx.com