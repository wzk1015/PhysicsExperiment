# 基础物理实验通用型数据处理程序

下载后通过main.exe运行<br>
最上面英文的提示无视就好（是python第三方库自己的问题）

1、密里根油滴实验（仅针对2020年春季线上仿真实验）：<br>
--
在`MillikanOilDrop\MillikanOilDropData.xlsx`里按照行与列的标签添加实验数据。<br>
在`MillikanOilDrop\MillikanOilDrop.txt`中修改实验所需的参数，原文件中数据、符号之间的空格请务必保留。<br>

* 数据中第一列是电压，后面八列是时间：<br>
* 第一、四、七……行是静态法中的下降时间。<br>
* 第二、五、八……行是动态法中的上升时间（这里的电压是动态法的平衡电压）。<br>
* 第三、六、九……行是动态法中的下降时间（这几行的电压没用）。<br>

打开`main.exe`后输入1即可得到结果。<br>



## 2、太阳能电池特性测量

* 将`SolarBattery/in.xlsx`中蓝底单元格填好
* 打开`main.exe`输入2，或进入SolarBattery目录直接python运行`SolarBattery.py`
* 结果输出至``SolarBattery/out.xlsx`

by wzk



助大家实验顺利！
---
