from docx import *
from tkinter import *
from datetime import datetime
from datetime import date

#todo
#rewrite as class
#in-house vs external version
#can we use lists and zips to create multiple jobs? each item i.e. jobtitle gets its own list, then these are zipped into job lists? 
#may make more sense to create directly into ind lists @create screen.

#Style
colorPrimary = '#AD2623'
colorSecondary = '#003352'
#End Style

#Window Rules
winMain = Tk()
winMain.title("Onesheet Generator")
winMain.geometry('800x600')
#Title
lblTitle = Label(winMain, text="Onesheet Generator", font=("Arial Bold", 25))
lblTitle.grid(column=1, row=0)

class addJob:
    def __init__(self, jobTitle, client, orderDate, startDate, payRate, heavyLifter, shift, hours, location, supervisor, openings, description, education, skills, experience, certifications):
        self.jobTitle = jobTitle
        self.client = client
        self.orderDate = orderDate
        self.startDate = startDate
        self.payRate = payRate
        self.heavyLifter = heavyLifter
        self.shift = shift
        self.hours = hours
        self.location = location
        self.supervisor = supervisor
        self.openings = openings
        self.description = description
        self.education = education
        self.skills = skills
        self.experience = experience
        self.certifications = certifications

    def clkAddJob():
        winAddJob = Toplevel()
        Label(winAddJob, text="Job Title").grid(column = 0, row = 2)
        Name = Entry(winAddJob).grid(column = 1, row = 2, columnspan = 3)

        Label(winAddJob, text="Client").grid(column = 0, row = 3)
        Client = Entry(winAddJob).grid(column = 1, row = 3, columnspan = 3)

        Label(winAddJob, text="Order Date").grid(column = 0, row = 4)
        OrderDate = Entry(winAddJob).grid(column = 1, row = 4, columnspan = 3)

        Label(winAddJob, text="Start Date").grid(column = 0, row = 5)
        StartDate = Entry(winAddJob).grid(column = 1, row = 5, columnspan = 3)

        Label(winAddJob, text="Pay Rate").grid (column = 0, row = 6)
        PayRate = Entry(winAddJob).grid(column = 1, row = 6, columnspan = 3)
        

        Label(winAddJob, text="Heavy Lifter?").grid (column = 0, row = 7)
        HeavyLifter = Radiobutton(winAddJob, text="yes", value="yes").grid (column = 1, row = 7)
        HeavyLifterNo = Radiobutton(winAddJob, text="no", value="no").grid (column = 3, row = 7)

        Label(winAddJob, text="Shift").grid (column = 0, row = 8)
        Shift1 = Radiobutton(winAddJob, text="1", value="1").grid (column = 1, row = 8)
        Shift2 = Radiobutton(winAddJob, text="2", value="2").grid (column = 2, row = 8)
        Shift3 = Radiobutton(winAddJob, text="3", value="3").grid (column = 3, row = 8)

        Label(winAddJob, text="Hours").grid (column = 0, row = 9)
        Hours = Entry(winAddJob).grid(column = 1, row = 9, columnspan = 3)

        Label(winAddJob, text="Location").grid (column = 0, row = 10)
        Location = Entry(winAddJob).grid(column = 1, row = 10, columnspan = 3)

        Label(winAddJob, text="Supervisor").grid (column = 0, row = 11)
        Supervisor = Entry(winAddJob).grid(column = 1, row = 11, columnspan = 3)

        Label(winAddJob, text="Number of Openings").grid (column = 0, row = 12)
        Openings = Entry(winAddJob).grid(column = 1, row = 12, columnspan = 3)

        #SPACERS
        Label(winAddJob, text="   ").grid(column = 4, row=1)
        Label(winAddJob, text="   ").grid(column = 6, row=1)
        Label(winAddJob, text="").grid(column = 5, row = 13)
        Label(winAddJob, text="").grid(column = 0, row = 1)
        #ENDSPACERS
        Label(winAddJob, text="Job Description").grid (column = 5, row = 1)
        JobDescription = Text(winAddJob, width = 30, height = 30).grid(column = 5, row = 2, rowspan = 11)

        Label(winAddJob, text="Education").grid (column = 7, row = 1)
        Education = Text(winAddJob, width = 30, height = 5).grid(column = 7, row = 2, rowspan = 2)

        Label(winAddJob, text="Experience").grid (column = 7, row = 4)
        Experience = Text(winAddJob, width = 30, height = 5).grid(column = 7, row = 5, rowspan = 2)

        Label(winAddJob, text="Skills").grid (column = 7, row = 7)
        Skills = Text(winAddJob, width = 30, height = 5).grid(column = 7, row = 8, rowspan = 2)

        Label(winAddJob, text="Certifications").grid (column = 7, row = 10)
        Certifications = Text(winAddJob, width = 30, height = 5).grid(column = 7, row = 11, rowspan = 2)

        def clkGenerateOnesheet(jobTitle, client, orderDate, startDate, payRate, heavyLifter, shift, hours, location, supervisor, openings, description, education, skills, experience, certifications):
            


            onesheet = Document()

            onesheet.add_heading(jobTitle, 0)

            #pulls current datetime, converts to string to use in filename (prevents duplicate files if loading multiple jobs)
            
            today = date.today()
            onesheet.save(str(client)+str(jobTitle)+str(today)+'.docx')

        Button(winAddJob, text="Generate Onesheet", height = 2, width = 20, bg=colorPrimary, fg="white", command=clkGenerateOnesheet(Name, Client, OrderDate, StartDate, PayRate, HeavyLifter, Shift1, Hours, Location, Supervisor, Openings, JobDescription, Education, Experience, Skills, Certifications)).grid (column = 5, row = 14)

        
        winAddJob.title("Add New Job")
        winAddJob.geometry('800x600')

    

    
    btnAddJob = Button(winMain, text="Add Job", bg=colorPrimary, fg="white", command=clkAddJob)
    btnAddJob.grid(column=0, row=3)
    btnAddJob.grid()

winMain.mainloop()