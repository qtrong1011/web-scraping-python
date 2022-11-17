from lmsJobs import CreateStudent, DriverBuilder



##### MAIN DRIVER #########
if __name__ == '__main__':
    driver = DriverBuilder().get_driver()
    student_info = [
            {
                "ems_id" : "ITSE0090",
                "center_code" : "ALAB-TEST",
                "student_name": "TEST-TL",
                "dob" : "19951110",
                "parent_name": ""
            },
            {
                "ems_id" : "ITSE0091",
                "center_code" : "QT",
                "student_name": "TEST-TL1",
                "dob" : "19951111",
                "parent_name": ""
            }
    ]
    create_student = CreateStudent()
    message = create_student.create_students(driver,student_info)
    print(message)