# 网易云音乐歌单导出Excel
一个简单的将网易云Web端歌单JSON导出成Excel的小工具

## 使用方法
打开需要导出的歌单，一般网址是  
https://music.163.com/#/my/m/music/playlist?id=XXXXXXX  
XXXXXXX是歌单ID

F12打开网络监测，并刷新页面，找到一个名叫detail的请求，请求地址一般是  
https://music.163.com/weapi/v3/playlist/detail  
网址后面可能有一些其他参数

复制完整的Response（应该是一整段JSON）至窗口中，并导出，程序会在当前目录下生成out.xlsx
Excel中四列分别是：歌名/歌手/专辑/时长
