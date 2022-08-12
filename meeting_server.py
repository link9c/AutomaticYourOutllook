from flask import Flask, jsonify, request, make_response, abort
import meetingMailApi


app = Flask(__name__)
app.debug = False

MY_URL = ''


# post
@app.route('/meeting_server/', methods=['POST'])
def post_task():
    mail = meetingMailApi.MeetingMail()
    # print(request)
    # print(request.json)
    if not request.json:
        abort(404)
    result = mail.json_rec(request.json)
    print(result)
    # print(jsonify(result))
    return jsonify(result)


# 404处理
@app.errorhandler(404)
def not_found(error):
    return make_response(jsonify({'error': 'Not found'}), 404)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8001)
