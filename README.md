获取OneDrive直链
===========
本应用的目的是为了批量获取OneDrive直链，方便在其他站点外链。前身为[提取OneDrive直链地址工具](http://bbs.comicdd.com/thread-354826-1-1.html)的网页版，因为原来的工具失效了，软件版也不是那么好用了，因此决定使用OneDrive官方API来进行获取，确保不失效。 

# 马上使用

https://mapaler.github.io/GetOneDriveDirectLink/

# 隐私声明

使用微软官方[OneDrive file picker SDK](https://dev.onedrive.com/sdk/js-v7/js-picker-overview.htm)，本应用不会得到你的账号密码和其他用户资料。
目前仅申请了Files.Read、Files.Read.Selected两个权限，API只会返回用户选择的文件的信息，不会获得其他内容。 