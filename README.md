# 健康码信息提取
批量处理健康码截图

### Install
`pip install -r requirements.txt`

### Using
`python main.py`

自带可视化界面，能够批量读取文件夹内所有截图，然后将数据提取到excel中。

可用 nuitka 打包，

`nuitka --windows-disable-console  main.py`

exe生成文件较大，需要可以直接找我 geeekhao@foxmail.com

第一次使用程序时，点击config.bat ， 会在C盘配置python库以及paddleocr模型，大小在1.5G

之后直接点击main.exe即可执行。