import email
import smtplib
import datetime
import icalendar
import uuid

from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart


class MeetingMail(object):
    def __init__(self):
        # 初始化事件，日历，邮件信息
        self.cal = icalendar.Calendar()
        self.event = icalendar.Event()
        self.msg = MIMEMultipart("alternative")

    def set_event(self, **kwargs):
        """
        :return: mail_uid,event
        """
        status = kwargs.get('status')
        self.event.add('status', kwargs.get('status') or "confirmed")
        self.event.add('attendee', kwargs.get('receiver'))
        self.event.add('dtstart', kwargs.get('start'))
        self.event.add('dtend', kwargs.get('end'))
        self.event.add('summary', kwargs.get('subj'))

        self.event.add('location', kwargs.get('location'))

        self.event.add('sequence', kwargs.get('sequence'))
        uid = kwargs.get('uid')
        if uid:
            self.event.add('uid', uid)
        else:
            self.event.add('uid', uuid.uuid1())
        if status:
            return self.event['uid'], self.event
        self.event.add('category', "Event")
        self.event.add('organizer', kwargs.get('sender'))
        self.event.add('description', kwargs.get('description'))
        self.event.add('created', datetime.datetime.now())

        return self.event['uid'], self.event

    def set_calendar(self, method, event):
        self.cal.add('prodid', '-//ACME/DesktopCalendar//EN')
        self.cal.add('version', '2.0')
        self.cal.add('method', method)
        self.cal.add_component(event)

        return self.cal

    @classmethod
    def set_part(cls, cal, method="REQUEST"):
        filename = "meeting.ics"
        part = MIMEBase('text', "calendar", method=method, name=filename)
        part.set_payload(cal.to_ical())
        email.encoders.encode_base64(part)
        part.add_header('Content-Description', filename)
        part.add_header("Content-class", "urn:content-classes:calendarmessage")
        part.add_header("Filename", filename)
        part.add_header("Path", filename)

        return part

    def set_msg(self, subj, sender, receiver, description, part):
        self.msg["Content-class"] = "urn:content-classes:calendarmessage"
        self.msg["Subject"] = subj
        self.msg["From"] = sender
        receiver = (',').join(receiver)
        self.msg["To"] = receiver
        self.msg.attach(MIMEText(description))
        # self.msg.del_param(part)
        self.msg.attach(part)

        return self.msg

    @classmethod
    def send_mail(cls, sender, receiver, msg):
        host = 'mail.microport.com.cn'
        port = 25
        smtp = smtplib.SMTP(host, port)
        try:
            smtp.sendmail(sender, receiver, msg.as_string())
            print('邮件成功发送至%s' % receiver)
            smtp.quit()
            smtp.close()
            return 1
        except Exception:
            print('发送邮件失败:%s' % receiver)
            return 0

    def send_appointment(self, **kwargs):
        """
        start, end, sender, receiver, subj, description, location
        """
        mail_uid, event = self.set_event(**kwargs, sequence=0)
        print(mail_uid)
        cal = self.set_calendar(method='REQUEST', event=event)
        part = self.set_part(cal)
        msg = self.set_msg(kwargs['subj'], kwargs['sender'], kwargs['receiver'], kwargs['description'], part)

        return mail_uid, msg

    def update_appointment(self, **kwargs):
        """
        start, end, sender, receiver, subj, description, location，uid
        """
        mail_uid, event = self.set_event(**kwargs, sequence=1)
        cal = self.set_calendar(method='REQUEST', event=event)
        part = self.set_part(cal)
        msg = self.set_msg(kwargs['subj'], kwargs['sender'], kwargs['receiver'], kwargs['description'], part)
        return mail_uid, msg

    def cancel_appointment(self, **kwargs):
        """
        start, end, sender, receiver, subj, description, location，uid
        """
        mail_uid, event = self.set_event(**kwargs, sequence=0, status='CANCELLED')
        cal = self.set_calendar(method='CANCEL', event=event)
        part = self.set_part(cal=cal, method='CANCEL')
        # print(kwargs)
        msg = self.set_msg(kwargs['subj'], kwargs['sender'], kwargs['receiver'], kwargs['description'], part)
        return mail_uid, msg

    def json_rec(self, json):  # (start, end, sender, receiver, subj, description, location, uid, operation)
        start = json['start']
        end = json['end']
        sender = json['sender']
        receiver = json['receiver']
        subj = json['subj']
        description = json['description']
        location = json['location']
        uid = json.get('uid')
        operation = json.get('operation')
        opt = None
        if operation == 'send':
            opt = self.send_appointment
        if operation == 'update':
            subj = '会议更新:%s' % subj
            opt = self.update_appointment
        if operation == 'cancel':
            subj = '已取消:%s' % subj
            opt = self.cancel_appointment
        if opt is None:
            return {'result': 0}

        mail_uid, msg = opt(start=datetime.datetime.strptime(start, "%Y-%m-%d %H:%M") - datetime.timedelta(hours=8)
                            , end=datetime.datetime.strptime(end, "%Y-%m-%d %H:%M") - datetime.timedelta(hours=8)
                            , sender=sender
                            , receiver=receiver
                            , subj=subj
                            , description=description
                            , location=location
                            , uid=uid)

        send = self.send_mail(sender, receiver, msg)
        if send == 1:
            return {'result': 1, 'uid': mail_uid}
        else:
            return {'result': 0}


if __name__ == '__main__':
    mail = MeetingMail()
    # mail.json_rec({"start": "2019-11-30 9:00", "end": "2019-11-30 10:00", "sender": "chenyulei@microport.com",
    #                "receiver": ['chenyulei@microport.com'],
    #                "subj": "会议测试邮件,修复了组织者,可以进行应答", "description": "qweqeq",
    #                "location": "xxxxx7F会议室"，'operation': 'send'})
    # mail.json_rec({"start": "2019-11-30 10:00", "end": "2019-11-30 11:00", "sender": "chenyulei@microport.com",
    #                "receiver": ['chenyulei@microport.com'],
    #                "subj": "会议测试邮件,修复了组织者,可以进行应答", "description": "qweqeq",
    #                "location": "xxxxx7F会议室", 'uid': 'fd973b4c-111c-11ea-9671-005056c00008'，'operation': 'update'})

    # mail.json_rec({"start": "2019-11-30 10:00", "end": "2019-11-30 11:00", "sender": "chenyulei@microport.com",
    #                "receiver": ['chenyulei@microport.com'],
    #                "subj": "会议测试邮件,修复了组织者,可以进行应答", "description": "qweqeq",
    #                "location": "xxxxx7F会议室", 'uid': 'fd973b4c-111c-11ea-9671-005056c00008', 'operation': 'cancel'})
    mail.json_rec({"start": "2019-12-06 14:00", "end": "2019-12-06 14:29", "sender": "mjwei@microport.com",
                   "receiver": ["chenyulei@microport.com"], "subj": "协同办公小组周会",
                   "description": "你好,\r\n 我们将在2019-12-06 11:00,于7F信息管理部小项目室召开协同办公小组周会", "location": "7F信息管理部小项目室",
                   "operation": "send", "uid": "a5e20d50-17e3-11ea-bbbb-4023437e7d6a"})
