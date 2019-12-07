
参数:
- start:开始时间
- end:结束时间
- sender:发送人
- receiver:接收人
- subj:标题
- description:邮件内容
- location:会议地址
- uid:邮件唯一标识符，update,cancel操作必须传入send返回的uid
- operation:send(发送会议邀请)，update(发送会议更新邮件)，cancel(关闭会议)

示例：
```python
# send操作时可不传uid
{"start": "2019-11-30 10:00", 
"end": "2019-11-30 11:00", 
"sender": "xxx@xxx.com",
"receiver": ["xxx@xxx.com"],
"subj": "会议测试邮件,修复了组织者,可以进行应答", 
"description": "欢迎参加会议！",
"location": "xxxxx7F会议室", 
"uid": "fd973b4c-111c-11ea-9671-005056c00008", 
"operation": "cancel"}
```

返回示例:
```python
# 邮件发送成功
{"result":1,"uid":"fd973b4c-111c-11ea-9671-005056c00008"}
# 邮件发送失败
{"result":0}
# 未传入参数
{"error":"Not found"}
```
