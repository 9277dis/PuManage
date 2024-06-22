# PU助手（Python）

## 开发环境：

### 		Win11、Pycharm专业版（2023.3）、python3.10（解释器）

## 下载地址：

| 名称            | 地址                                                         | 备注                                                         |
| :-------------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| pycharm         | [JetBrains: Essential tools for software developers and teams](https://www.jetbrains.com/) | 官网下载即可，建议2023版本                                   |
| python3.10      | https://www.python.org/downloads/release/python-3100/        | 这是环境解释器，直接安装即可，默认安装会自动被检测为系统环境解释器 |
| pycharm激活工具 | http://t.0cs.cc/                                             | 一键激活，但只支持win版本                                    |

## 其他

#### 遇到这种报错![image-20240622183803908](C:\Users\9277\AppData\Roaming\Typora\typora-user-images\image-20240622183803908.png)

#### `Alt+Enter`，选择`安装`即可

#### 打包：

```shell
PyInstaller --onefile --noconsole --hidden-import=tkinter --name "pu助手1.08" --icon "pu.ico" --add-data "pu.ico;."  main.py
```

