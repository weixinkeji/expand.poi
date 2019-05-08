 需要的依赖
 
 ```
 <!-- xls -->
<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi</artifactId>
	<version>4.1.0</version>
</dependency>
<!-- xlsx读写 -->
<dependency>
	<groupId>org.apache.poi</groupId>
	<artifactId>poi-ooxml</artifactId>
	<version>4.1.0</version>
</dependency>
<!-- xlsx依赖需要 -->
<dependency>
	<groupId>org.dom4j</groupId>
	<artifactId>dom4j</artifactId>
	<version>2.1.1</version>
</dependency>
```

```
//一键导出excel文档
OutputStream out= new FileOutputStream("D:\\resources\\学生信息.xlsx");
JWEOfficeW.writeToExcel_xlsx(out,"aaaaaaaa", list);

//一键导入文档
List<Entity> list =JWEOfficeR.readExcel_xlsx(Entity.class, "aaaaaaaa", "D:\\resources\\学生信息.xlsx");


//复杂应用，一次写多个工作表
	JWEOfficeW obj=new JWEOfficeW("D:\\resources\\学生信息.xlsx",JWEOfficeEnum.xlsx);
			 obj.addToExcel("aaaaaaaa", list);
			 obj.addToExcel("aaaaaaaa2", list);
			 obj.writeAndAutoCloseIO();//关闭


//一次读多个工作表的数据
JWEOfficeR obj=new JWEOfficeR("D:\\resources\\学生信息.xlsx",JWEOfficeEnum.xlsx);
		List<Entity> list =obj.readExcel(Entity.class, "aaaaaaaa");
		for(Entity e:list) {
			System.out.println(e.getId()+",age="+e.getAge()+",d="+e.getD()+",name="+e.getName()+",BirthDay="+e.getBirthDay()+",Time="+e.getTime());
		}
		
		list =obj.readExcel(Entity.class, "aaaaaaaa2");
		for(Entity e:list) {
			System.out.println(e.getId()+",age="+e.getAge()+",d="+e.getD()+",name="+e.getName()+",BirthDay="+e.getBirthDay()+",Time="+e.getTime());
		}
		
		

更功能功能，请关注源码


		 
```

以上方法的代码，需要实体类配合注解，@JWEOffice 的使用
示例如下：

```
public class Entity {
	// @JWEOffice(title = "主键", sort = 1)
	private Integer id;
	@JWEOffice(title = "名称", sort = 8)
	private String name;
	@JWEOffice(title = "年龄", sort = 4)
	private int age;
	@JWEOffice
	private double d;
	@JWEOffice(title = "生日3", sort = 3, dateformat = "yyyy年MM月dd日")
	private Date birthDay;
	@JWEOffice(title = "时辰5", sort = 5, dateformat = "hh:mm:ss")
	private Date time;
  
//title：excel文档，第一行的标题   
//sort：全部属性（title）在excel文档的列中，按sort值，从小到大排序
//dateformat：表示显示格式。仅对时间有用
  
```




