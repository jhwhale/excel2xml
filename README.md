# excel2xml
将excel文件中的每一张工作表转为一个xml文件，文件名称为对应工作表的名称。
每一行为一条case
需在第一行标明每列名称
	第一列为case标题：name
	第二列为case前提条件：preconditions
	第三列为执行步骤：steps
	第四列为期望结果：expectedresults

生成的xml文件保存在excel文件所在目录下
