微软outlook邮箱 会议邀请邮件api服务


请求URL:

    /meeting_server/
    
请求方式:

    POST
参数:
- start:开始时间
- end:结束时间
- sender:发送人
- required:必要参加人
- optional:可选参加人
- receiver:{required:必须,optional:可选}
- subj:标题
- description:邮件内容
- location:会议地址
- uid:邮件唯一标识符，update,cancel操作必须传入send返回的uid
- operation:send(发送会议邀请)，update(发送会议更新邮件)，cancel(关闭会议)
- repeat:重复周期参数

单个邮件示例：
```python
# send操作时可不传uid
{"start": "2019-11-30 10:00", 
"end": "2019-11-30 11:00", 
"sender": "xxx@xxx.com",       
"receiver": {
"required":["xxx@xxx.com"],
"optional":["xxx@xxx.com"]
},        
"subj": "会议测试邮件,修复了组织者,可以进行应答", 
"description": "欢迎参加会议！",    
"location": "xxxxx7F会议室", 
"uid": "fd973b4c-111c-11ea-9671-005056c00008", 
"operation": "cancel"}
```
重复邮件示例：
```python
# send操作时可不传uid
{"start": "2019-11-30 10:00", 
"end": "2019-11-30 11:00", 
"sender": "xxx@xxx.com",       
"receiver": {
"required":["xxx@xxx.com"],
"optional":["xxx@xxx.com"]
},        
"subj": "会议测试邮件,修复了组织者,可以进行应答", 
"description": "欢迎参加会议！",    
"location": "xxxxx7F会议室", 
"uid": "fd973b4c-111c-11ea-9671-005056c00008", 
"operation": "cancel",
"repeat": {"FREQ": "WEEKLY", "COUNT":3, "BYDAY": "MO"}
}
```
```python
"repeat": {"FREQ": "WEEKLY", "COUNT":3, "BYDAY": "MO"}
"start": "2019-11-30 10:00"
"end": "2019-11-30 11:00"
```
这里repeat字段代表：每周周一发送，重复3次

开始和结束时间从start，end取

开始时间:10:00

结束时间:11:00

开始循环时间:2019-11-30

结束循环时间:2019-11-30后第三周的周一

其他示例:
- FREQ=DAILY;COUNT=10;INTERVAL=3 隔3天发送 重复10次
- FREQ=WEEKLY;COUNT=1;BYDAY=[MO,TU,WE,TH,FR] 工作日发送邮件，重复一次
- FREQ=MONTHLY;COUNT=10;BYMONTHDAY=3 每月第三天，重复10次
- FREQ=MONTHLY;COUNT=10;BYMONTHDAY=3;INTERVAL=3 每3个月，这个月第三天，重复10次
- FREQ=MONTHLY;COUNT=10;BYDAY=WE;BYSETPOS=2 每个月第二周的星期三，重复10次
- FREQ=YEARLY;COUNT=1;BYMONTHDAY=12;BYMONTH=8 每年8月12号，重复1次
- FREQ=YEARLY;COUNT=1;BYDAY=WE;BYMONTH=8;BYSETPOS=2 每年8月的第二个星期三，重复1次


返回示例:
```python
# 邮件发送成功
{"result":1,"uid":"fd973b4c-111c-11ea-9671-005056c00008"}
# 邮件发送失败
{"result":0}
# 未传入参数
{"error":"Not found"}
```

