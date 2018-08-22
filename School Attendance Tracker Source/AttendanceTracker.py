import xml.etree.ElementTree as ET
import csv
import xlsxwriter
import sys



if
infile = "cumulative_attendance.xml" 
file_xml = infile

tree = ET.parse(file_xml)
root = tree.getroot()

attendance_threshold = 90.0

show_missing_students = False;

def StudentOffRoll(students, search_id):
    try:
        stu_index = students.index(search_id)
        return stu_index;
    except ValueError:
        return None;

def StudentExists(students, search_id):
    for s in students:
        if(search_id == s.student_id):
            return s
        
    return None

class Attendance:
    def __init__(self):
        self.percent = ''
        self.startdate = ''
        self.enddate = ''
 
class Student:
    
    def __init__(self):
        self.student_id=0 
        self.name='' #full name
        self.chosenname = '' #first name
        self.form = ''
        self.tutor=''
        self.attendance=[] # list of all the students' attendance
        self.curr_attendance_week=0
        self.average_attendance = 0
        
          


#update with lates tutors
tutors = {'12SC':'Ms. Carr',
          '12SCG':'Ms. Gardiner', 
          '12KXM': 'Ms. Muir',
          '12KP' : 'Mr. Patel',
          '12JMO' : 'Mr. Oddy',
          '12JRD' : 'Mr. Dennett',
          '12SFM' : 'Ms. Munday',
          '12AS': 'Mr. Shouber',
          '12XX' : 'XX'
          }




params = root.findall("Parameter")
 
week_counter = 0
week_dates = []
all_students = []
all_off_roll_students = []

off_roll_student = []
 
with open('CSV_OffRoll.csv', 'r') as csvfile:
    data = csv.reader(csvfile, delimiter=',')
    for row in data:
        for cell in row:
            off_roll_student.append(cell)
        stu_id = off_roll_student[1] + off_roll_student[0] #index one contains the form, index 0 contains the full name
        all_off_roll_students.append(stu_id)
        off_roll_student.clear()

print("all off roll ", all_off_roll_students)


#for each week
for weeks in root.iter('Header'):

    #get the start and end date of the report
    for params in weeks.iter('Parameters'):
        for dates in params.iter('Period'):
            sdate = dates.findtext('Start')
            edate = dates.findtext('End')

    print(week_counter , ': ' , sdate, '-', edate)

    week_dates.append( str(sdate + '-' + edate) )

    
    #for each tutor base
    for forms in weeks.iter('Group'):
        form = forms.findtext('GroupName')
        #print(form)
        #for each student
        for pupils in forms.iter('PupilInformation'):
            pupil = pupils.findtext('FullName')

            if ((form not in tutors) == True):
                print('not adding form ', form)
                continue
            else:
                tutor = tutors[form]

                #work out the first name
                fname = pupil.split(",")
                ncount = len(fname)
                chosen_name = fname[ncount-1]
                
                full_name = pupil

                #work out attendance
                for am_reg in  pupils.iter('PupilAMMarks'): #am registration
                    am_reg_unath_percent = am_reg.findtext('UnauthorisedAbsences')
                for pm_reg in  pupils.iter('PupilPMMarks'): #pm registration
                    pm_reg_unath_percent = pm_reg.findtext('UnauthorisedAbsences')

                #calculate the unautharised percent
                unauth_percent = float(((float(am_reg_unath_percent) + float(pm_reg_unath_percent)) / 2)) #average am and pm
                final_unauth = 100 - round(unauth_percent)
                
                if(week_counter == 0):
                    print(form + " " +full_name + str(final_unauth)+ "%"  )

                stu_id = form + full_name

                stu = StudentOffRoll (all_off_roll_students,stu_id)
                if(stu != None): # do not add students that are off roll
                    #print("NOT ADDING OFF ROLL STUDENT: ", stu_id)
                    continue
                stu = StudentExists (all_students,stu_id)


                if(stu == None):

                    s = Student() #add the student to the list

                    if(week_counter != 0): # should not be adding students after the latest week
                       # print('WARNING: adding student after W0 ', week_counter, stu_id, tutor, chosen_name)
                        b = Attendance()
                        b.percent = -1
                        b.startdate = 0
                        b.enddate = b
                        #buffer the attendences, create a sentinal value for missing attendances
                        if(show_missing_students == True):
                            for x in range(0,week_counter):
                                s.attendance.insert(x,b)
                        else:
                            continue;
                
                    s.student_id = stu_id
                    s.name = full_name
                    s.chosenname = chosen_name
                    s.form = form
                    s.tutor = tutor
                    a = Attendance()
                    a.percent = final_unauth
                    a.startdate = sdate
                    a.enddate = edate
                    s.attendance.insert(week_counter,a)
                    
                    all_students.append(s)
                    
                     
                else:
                    a = Attendance() #append the students' attendance list
                    a.percent = final_unauth
                    a.startdate = sdate
                    a.enddate = edate
     
                    stu.attendance.insert(week_counter,a)
                       

    week_counter = week_counter + 1        


print("number of students is ", len(all_students))


 
# Create a workbook and add a worksheet.
workbook = xlsxwriter.Workbook('y12_attendance.xlsx')
worksheet = workbook.add_worksheet()

workbook_tutors = xlsxwriter.Workbook('y12_attendance_tutors.xlsx')
worksheet_tutors = workbook_tutors.add_worksheet()

bold = workbook.add_format({'bold': True})
boldT = workbook_tutors.add_format({'bold': True})

# Start from the first cell. Rows and columns are zero indexed.
row = 4
col = 0

excel_data2 = ['Full Name','Form','Average']
for w in week_dates:
    excel_data2.append(w) #add all the week dates to the top row

#write the header information
for item in excel_data2:
    worksheet.write(row, col, item, bold)
    if(col < 4):
        worksheet_tutors.write(row, col, item,boldT)  
    col += 1
    

row += 1 #next row


#cell background colours
green_format = workbook.add_format()
green_format.set_pattern(1)  
green_format.set_bg_color('#99CC00')

amber_format = workbook.add_format()
amber_format.set_pattern(1)   
amber_format.set_bg_color('#ff9900')

red_format = workbook.add_format()
red_format.set_pattern(1)   
red_format.set_bg_color('#ff0000')


#cell background colours for tutors
tgreen_format = workbook_tutors.add_format()
tgreen_format.set_pattern(1)  
tgreen_format.set_bg_color('#99CC00')

tamber_format = workbook_tutors.add_format()
tamber_format.set_pattern(1)   
tamber_format.set_bg_color('#ff9900')

tred_format = workbook_tutors.add_format()
tred_format.set_pattern(1)   
tred_format.set_bg_color('#ff0000')



#write attendance to excel file
with open('CSV_year12Attendance.csv', 'w', newline='') as fp:
    a = csv.writer(fp, delimiter=',')

    #write the header rows
    excel_data = ['Full Name','Form','Average']
    for w in week_dates:
        excel_data.append(w)

    data = [excel_data]
    a.writerows(data)
    

    #write the student information
    excel_data.clear();

    col =  0

    #loop through all students
    for k in all_students:

        #read the last attendance
        latest_attendance = k.attendance[0].percent

        num_attendances = 0
        sum_attendance = 0
        average_attendance = 0

        #check if the last attendance is less than the tresh hold
        if ((float(latest_attendance)) < attendance_threshold ):

            #write "name" "form", "attendances" to list
            excel_data = [k.name,k.form]
            
            for j in k.attendance:
                excel_data.append(j.percent)
                
                sum_attendance = sum_attendance + j.percent #work out the mean
                num_attendances = num_attendances + 1
                  
            #add the mean
            average_attendance = sum_attendance / num_attendances
            excel_data.insert(2,int(average_attendance))
         
            #write the student rows to the excel CSV file
            stu_data = [excel_data]
            a.writerows(stu_data)

            #write the xls file with coloured backgrounds
            for dat in excel_data:

                #get the current attendance
                if(col > 1 ): # index at 0 and 1 is name and form
                    #set the background colour
                    percentage = int(dat)
                    if(percentage <= 50):
                        worksheet.write(row, col, dat,red_format)
                    elif(percentage > 50 and percentage < 79):
                        worksheet.write(row, col, dat,amber_format)
                    else:
                        worksheet.write(row, col, dat,green_format)
                else:
                    worksheet.write(row, col, dat)
                    worksheet_tutors.write(row, col, dat)

                if(col > 1 and col < 4 ):
                    #set the background colour
                    if(percentage <= 50):
                        worksheet_tutors.write(row, col, dat, tred_format)
                    elif(percentage > 50 and percentage < 79):
                        worksheet_tutors.write(row, col, dat, tamber_format)
                    else:
                        worksheet_tutors.write(row, col, dat, tgreen_format)
                #else:
                    #worksheet_tutors.write(row, col, dat)
                
                    
                col +=  1 #next attendance
            row += 1 #next student
            col =  0 #next student 
            


#write mail merge letter values
with open('CSV_year12AttendanceLetters.csv', 'w', newline='') as fpe:
    l = csv.writer(fpe, delimiter=',')

    excel_data = ['Full Name','ChosenName','Attendance','Form','Tutor','StartDate','EndDate']
    data = [excel_data]
    l.writerows(data)
 
    excel_data.clear()  
    for c in all_students:
        latest_attendance = c.attendance[0].percent
        if ((float(latest_attendance)) < attendance_threshold ):
            excel_data = [c.name,c.chosenname,c.attendance[0].percent,c.form,c.tutor,c.attendance[0].startdate,c.attendance[0].enddate]
            fdata = [excel_data]
            l.writerows(fdata)



    
    

workbook.close()
workbook_tutors.close()

 
 

 










