from docx import *
from tkinter import *
from datetime import datetime
from datetime import date
from tkinter.messagebox import showerror

#Style
colorPrimary = '#AD2623'
colorSecondary = '#003352'
#End Style

#Window Rules
winMain = Tk()
winMain.title("Onesheet Generator")
winMain.geometry('1000x750')
#Title
lblTitle = Label(winMain, text="Onesheet Generator", font=("Arial Bold", 25))
lblTitle.grid(column=1, row=0)

Label(winMain, text="Job Title").grid(column = 0, row = 2)
jobTitle = Entry(winMain)
jobTitle.grid(column = 1, row = 2, columnspan = 3)

Label(winMain, text="Client").grid(column = 0, row = 3)
client = Entry(winMain)
client.grid(column = 1, row = 3, columnspan = 3)

Label(winMain, text="Order Date").grid(column = 0, row = 4)
orderDate = Entry(winMain)
orderDate.grid(column = 1, row = 4, columnspan = 3)

Label(winMain, text="Start Date").grid(column = 0, row = 5)
startDate = Entry(winMain)
startDate.grid(column = 1, row = 5, columnspan = 3)

Label(winMain, text="Pay Rate").grid (column = 0, row = 6)
payRate = Entry(winMain)
payRate.grid(column = 1, row = 6, columnspan = 3)

Label(winMain, text="Heavy Lifter?").grid (column = 0, row = 7)
heavyLifter = StringVar(winMain)
heavyLifterOptions = {"Yes": True, "No": False}
heavyLifterSelect = OptionMenu(winMain, heavyLifter, *heavyLifterOptions.keys())
heavyLifterSelect.grid (column = 1, row = 7)
#to get - heavyLifterState = heavyLifterOptions[heavyLifter.get()]

Label(winMain, text="Shift").grid (column = 0, row = 8)
shift = StringVar(winMain)
shiftOptions = {"First": 1, "Second": 2, "Third": 3}
shiftSelect = OptionMenu(winMain, shift, *shiftOptions.keys())
shiftSelect.grid (column = 1, row = 8)

Label(winMain, text="Hours").grid (column = 0, row = 9)
hours = Entry(winMain)
hours.grid(column = 1, row = 9, columnspan = 3)

Label(winMain, text="Location").grid (column = 0, row = 10)
location = Entry(winMain)
location.grid(column = 1, row = 10, columnspan = 3)

Label(winMain, text="Supervisor").grid (column = 0, row = 11)
supervisor = Entry(winMain)
supervisor.grid(column = 1, row = 11, columnspan = 3)

Label(winMain, text="Number of Openings").grid (column = 0, row = 12)
openings = Entry(winMain)
openings.grid(column = 1, row = 12, columnspan = 3)

#SPACERS
Label(winMain, text="   ").grid(column = 4, row=1)
Label(winMain, text="   ").grid(column = 6, row=1)
Label(winMain, text="").grid(column = 5, row = 13)
Label(winMain, text="").grid(column = 0, row = 1)
#ENDSPACERS

Label(winMain, text="Job Description").grid (column = 5, row = 1)
jobDescription = Text(winMain, width = 30, height = 30)
jobDescription.grid(column = 5, row = 2, rowspan = 11)

Label(winMain, text="Education").grid (column = 7, row = 1)
education = Text(winMain, width = 30, height = 5)
education.grid(column = 7, row = 2, rowspan = 2)

Label(winMain, text="Experience").grid (column = 7, row = 4)
experience = Text(winMain, width = 30, height = 5)
experience.grid(column = 7, row = 5, rowspan = 2)

Label(winMain, text="Skills").grid (column = 7, row = 7)
skills = Text(winMain, width = 30, height = 5)
skills.grid(column = 7, row = 8, rowspan = 2)

Label(winMain, text="Certifications").grid (column = 7, row = 10)
certifications = Text(winMain, width = 30, height = 5)
certifications.grid(column = 7, row = 11, rowspan = 2)

def clkGenerateOneSheet():
    jobTitleValue = jobTitle.get()
    clientValue = client.get()
    orderDateValue = orderDate.get()
    startDateValue = startDate.get()
    payRateValue = payRate.get() 
    if heavyLifterOptions[heavyLifter.get()] == True:
        heavyLifterDisplay = "Yes"
    elif heavyLifterOptions[heavyLifter.get()] == False:
        heavyLifterDisplay = "No"
    else:
        showerror(title = "Error", message = "You must select Yes or No for Heavy Lifter")
    shiftSelectValue = shiftOptions[shift.get()]
    hoursValue = hours.get()
    locationValue = location.get()
    supervisorValue = supervisor.get()
    openingsValue = openings.get()



    onesheet = Document()

    onesheet.add_heading(jobTitleValue + " @ " + clientValue, 0)

    #builds and display table
    displayTable = onesheet.add_table(rows=9, cols=2)
    
    label_cells = displayTable.columns[0].cells
    label_cells[0].text = 'Order Date'
    label_cells[1].text = 'Start Date'
    label_cells[2].text = 'Pay Rate'
    label_cells[3].text = 'Heavy Lifter'
    label_cells[4].text = 'Shift'
    label_cells[5].text = 'Hours'
    label_cells[6].text = 'Location'
    label_cells[7].text = 'Supervisor'
    label_cells[8].text = 'Openings'

    data_cells = displayTable.columns[1].cells
    data_cells[0].text = orderDateValue
    data_cells[1].text = startDateValue
    data_cells[2].text = payRateValue
    
    data_cells[3].text = heavyLifterDisplay
    data_cells[4].text = str(shiftSelectValue)
    data_cells[5].text = hoursValue
    data_cells[6].text = locationValue
    data_cells[7].text = supervisorValue
    data_cells[8].text = openingsValue


    today = date.today()
    onesheet.save(str(clientValue)+str(jobTitleValue)+str(today)+'.docx')

Button(winMain, text="Generate Onesheet", height = 2, width = 20, bg = colorPrimary, fg="white", command=clkGenerateOneSheet).grid (column = 5, row = 14)



winMain.mainloop()