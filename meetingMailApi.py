import email
import smtplib
import datetime
import icalendar
import uuid

from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart

from logger import Logger

log = Logger('Api')


class MeetingMail(object):
    def __init__(self):
        # 初始化事件，日历，邮件信息
        self.cal = icalendar.Calendar()
        # self.ics = icalendar.Event()
        icalendar.Journal
        self.msg = MIMEMultipart("alternative")
        self.ics = self.base_ics()
        print(1, self.ics)

    def base_ics(self):
        """
        读取ics文件
        """
        with open('base.ics', 'rb') as fp:
            gcal = icalendar.Calendar.from_ical(fp.read())
        return gcal

    def set_vevent(self, **kwargs):
        """
        :return: mail_uid,event
        """
        status = kwargs.get('status')
        receiver = kwargs.get('receiver')
        address = kwargs.get('address')
        # 必须参加人员
        req = receiver.get('required')
        # 可选参加人员
        opt = receiver.get('optional')
        # 设置参数
        date_ics = icalendar.prop.vDatetime
        # 参加人员
        self.ics['attendee'] = req
        # 时间
        st = kwargs.get('start')
        end = kwargs.get('end')

        # ----------------
        self.ics['dtstart'] = date_ics(st)

        self.ics['dtend'] = date_ics(end)
        self.ics.pop('dtstamp')
        self.ics.pop('last-modified')
        self.ics['created'] = datetime.datetime.now()
        # -----------------
        self.ics['summary'] = kwargs.get('subj')
        self.ics['location'] = kwargs.get('location')
        self.ics['sequence'] = kwargs.get('sequence')
        self.ics['organizer'] = kwargs.get('sender')
        self.ics['description'] = kwargs.get('description')
        repeat = kwargs.get('repeat')

        if repeat:
            # FREQ=DAILY;COUNT=10;INTERVAL=3 隔3天发送 重复10次
            # FREQ=WEEKLY;COUNT=1;BYDAY=MO,TU,WE,TH,FR 工作日发送邮件，重复一次
            # FREQ=MONTHLY;COUNT=10;BYMONTHDAY=3 每月第三天，重复10次
            # FREQ=MONTHLY;COUNT=10;BYDAY=WE;BYSETPOS=2 每个月第二周的星期三，重复10次
            # FREQ=YEARLY;COUNT=1;BYMONTHDAY=12;BYMONTH=8 每年8月12号，重复1次
            # RRULE:FREQ=YEARLY;COUNT=1;BYDAY=WE;BYMONTH=8;BYSETPOS=2 每年8月的第二个星期三，重复1次
            ruler = {'FREQ': 'DAILY', 'COUNT': 1}
            for key in ['FREQ', 'COUNT', 'INTERVAL', 'BYDAY', 'BYMONTHDAY', 'BYMONTH', 'BYSETPOS']:
                if repeat.get(key):
                    # print(type(repeat.get(key)))
                    ruler[key] = repeat.get(key)

            self.ics['rrule'] = icalendar.vRecur(ruler)
        else:
            self.ics.pop('rrule')

        # 新增参数
        # 状态
        self.ics.add('status', kwargs.get('status') or "confirmed")
        # 可选择人员
        if opt:
            self.ics.add('attendee;ROLE=OPT-PARTICIPANT', opt)

        if address:
            self.ics.add('attendee', icalendar.vCalAddress(address))
        # self.event.add('dtstart', '16010101T000000')

        uid = kwargs.get('uid')
        if uid:
            self.ics.add('uid', uid)
        else:
            self.ics.add('uid', uuid.uuid1())
        if status:
            return self.ics['uid'], self.ics

        print(self.ics)
        return self.ics['uid'], self.ics

    def set_calendar(self, method, event):

        self.cal.add('prodid', '-//Microsoft Corporation//Outlook 16.0 MIMEDIR//EN')
        self.cal.add('version', '2.0')
        self.cal.add('method', method)
        self.cal.add_component(event)
        # print(self.cal)
        return self.cal

    @classmethod
    def set_part(cls, cal, method="REQUEST"):
        filename = "meeting.ics"
        part = MIMEBase('text', "calendar", method=method, name=filename)

        part.set_payload(cal.to_ical())
        # print(cal.to_ical())
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
        req = receiver.get('required')
        opt = receiver.get('optional')

        if opt:
            rec = req + opt
        else:
            rec = req
        if rec:
            rec = ','.join(receiver)
        else:
            # rec = "zhoupeng@microport.com"
            print("没有recever也没有optional")

        self.msg["to"] = rec

        self.msg.attach(MIMEText(description, "HTML"))
        self.msg.attach(part)

        return self.msg

    @classmethod
    def send_mail(cls, sender, receiver, msg):

        host = '127.0.0.1'
        port = 25
        try:
            smtp = smtplib.SMTP(host, port)
        except Exception as e:
            # print('发送邮件失败:%s,error:%s' % (receiver, e))
            log.get_log().warn('发送邮件失败:%s,error:%s' % (receiver, e))
            return 0
        try:
            req = receiver.get('required')
            opt = receiver.get('optional')
            if opt:
                receiver = req + opt
            else:
                receiver = req
            smtp.sendmail(sender, receiver, msg.as_string())
            log.get_log().info('邮件成功发送至%s' % receiver)
            # print('邮件成功发送至%s' % receiver)
            smtp.quit()
            smtp.close()
            return 1
        except Exception as e:
            log.get_log().warn('发送邮件失败:%s,error:%s' % (receiver, e))
            return 0

    def send_appointment(self, **kwargs):
        """
        start, end, sender, receiver, subj, description, location
        """
        mail_uid, event = self.set_vevent(sequence=0, **kwargs)

        log.get_log().info(mail_uid)
        # print(mail_uid)
        cal = self.set_calendar(method='REQUEST', event=event)
        part = self.set_part(cal)
        msg = self.set_msg(kwargs['subj'], kwargs['sender'], kwargs['receiver'], kwargs['description'], part)

        return mail_uid, msg

    def update_appointment(self, **kwargs):
        """
        start, end, sender, receiver, subj, description, location，uid
        """
        mail_uid, event = self.set_vevent(sequence=1, **kwargs)
        cal = self.set_calendar(method='REQUEST', event=event)
        part = self.set_part(cal)
        msg = self.set_msg(kwargs['subj'], kwargs['sender'], kwargs['receiver'], kwargs['description'], part)
        return mail_uid, msg

    def cancel_appointment(self, **kwargs):
        """
        start, end, sender, receiver, subj, description, location，uid
        """
        mail_uid, event = self.set_vevent(sequence=0, status='CANCELLED', **kwargs)
        cal = self.set_calendar(method='CANCEL', event=event)
        part = self.set_part(cal=cal, method='CANCEL')
        # print(kwargs)
        msg = self.set_msg(kwargs['subj'], kwargs['sender'], kwargs['receiver'], kwargs['description'], part)
        return mail_uid, msg

    def json_rec(self, json):  # (start, end, sender, receiver, subj, description, location, uid, operation)
        try:
            start = json['start']
            end = json['end']
            sender = json['sender']
            receiver = json['receiver']
            subj = json['subj']
            description = json['description']
            location = json['location']

        except Exception as e:
            return {'result': 0, 'error': str(e)}
        try:
            address = json.get('address')
        except Exception as e:
            return {'result': 0, 'error': str(e)}

        if not receiver:
            return {'result': 0, 'error': '请选择接收人'}

        uid = json.get('uid')
        repeat = json.get('repeat')
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
        log.get_log().info(f"{json}")
        mail_uid, msg = opt(start=datetime.datetime.strptime(start, "%Y-%m-%d %H:%M") - datetime.timedelta(hours=8)
                            , end=datetime.datetime.strptime(end, "%Y-%m-%d %H:%M") - datetime.timedelta(hours=8)
                            , sender=sender
                            , receiver=receiver
                            , subj=subj
                            , description=description
                            , location=location
                            , uid=uid
                            , repeat=repeat
                            , address=address)

        send = self.send_mail(sender, receiver, msg)
        print(str(mail_uid))
        if send == 1:
            return {'result': 1, 'uid': str(mail_uid)}
        else:
            return {'result': 0}


if __name__ == '__main__':
    mail = MeetingMail()
   

    message = {
        "start": "2021-08-06 8:00",
        "end": "2021-08-6 11:00",
        "sender": "XX@xx.com",
        "address":"address@xx.com",
        "receiver": {
            "required": ["xx@xx.com"],
            "optional": ["xx@xx.com"]
        },
        "subj": "测试9",
        "description": "<p><span>您好，</span></p><p><span>周期会议：测试9请您准时参加！</span></p><p><span>日期：2020-08-03--2020-08-06 18:47</span></p><p><span>时间：18:45--18:47</span></p><p><span>地点：7F信息管理部小项目室</span></p><p><span>定期模式：每1天</span></p><p><span>会议内容：9999</span></p><p><span>此邮件由portal系统自动发出。</span></p>",
        "location": "7F",
        "operation": "send",
        "repeat": {
            "FREQ": "WEEKLY",
            "INTERVAL": 1,
            "COUNT": 2,
            "BYDAY": ["MO", "TU"]

        }

        
    }

    mail.json_rec(message)
