from asyncio.log import logger
from distutils.log import debug
# import sys
# sys.path.insert(0,"D:\web-scraping")
from selenium import webdriver
from flask import Flask , request, abort, jsonify
from lmsJobs import MovingCenterLMS, DriverBuilder, LoginAdminLMS


app = Flask(__name__)

@app.route('/LMS/api/movingCenter', methods=['POST'])
def moving_center():
    driver = DriverBuilder().get_driver()
    login = LoginAdminLMS().login_mag_admin(driver)
    moving_process = MovingCenterLMS()
    if not request.json or not 'MaHV' in request.json:
        abort(400)
    student_info = {
        'Name': request.json['Name'],
        'MaHV' : request.json['MaHV'],
        'CenterCurrent' : request.json['CenterCurrent'],
        'CenterMoving': request.json['CenterMoving']
    }
    moved_student = moving_process.moving_center_LMS(driver,student_info)
    return jsonify({'moved_student': moved_student}), 201
@app.route('/home')
def home():
    return "Hello Azure!!!!!"


if __name__ == '__main__':
    try:
        app.run()
    except Exception as e:
        logger.error('Failed to run the app: ' + str(e))