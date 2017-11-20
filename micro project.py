import xlrd
import random
import string
locat="R1 KTU Student data.xlsx"
workbook=xlrd.open_workbook(locat)
sheet=workbook.sheet_by_index(0)

def grade_to_marks(grade):
    grades=("O[S]","A+","A","B+","B","C","P","F","FE")
    gp=(10,9,8.5,8,7,6,5,0)
    marksrange=((89,100),(84,90),(79,85),(69,80),(59,70),(49,60),(44,50),(-1,45))
    if (grade=="FE"):
        pt=0
        mark="Failed due to eligibility criteria"
        return mark,pt
    else:
        for i in range(len(grades)):
            if (grade==grades[i]):
                pt=gp[i]
                mark=random.randint(marksrange[i][0],marksrange[i][1])
                return mark,pt

def code_to_sub(code):
    codes=("S1MA101","S1CY100","S1BE100","S1BE10105","S1BE103","S1EC100","S1CY110","S1CS110","S1EC110")
    subs=("CAL","CHE","EM","ICE","ISE","BE","CHE LAB","ICE WRKSHOP","BE WRKSHOP")
    for i in range(len(codes)):
        if (code==codes[i]):
            sub=subs[i]
            return sub

def sub_to_code(sub):
    codes=("S1MA101","S1CY100","S1BE100","S1BE10105","S1BE103","S1EC100","S1CY110","S1CS110","S1EC110")
    subs=("CAL","CHE","EM","ICE","ISE","BE","CHE LAB","ICE WRKSHOP","BE WRKSHOP")
    for i in range(len(codes)):
        if (sub==subs[i]):
            code=codes[i]
            return code

def stud_details():
    rno=raw_input("Enter the roll no. or register no.:")
    if (rno.isdigit()):
        for i in range(22):
            print (sheet.cell_value(0, i)),"      :      ",(sheet.cell_value(int(rno), i))
        for i in range(22,sheet.ncols):
            sub=code_to_sub(str(sheet.cell_value(0,i)))
            mark,pt=grade_to_marks(str(sheet.cell_value(int(rno),i)))
            print sub,"(",sheet.cell_value(0,i),")","     :     ",mark,"   ;   ",str(sheet.cell_value(int(rno),i)),"(",pt,")"
    else:
        for i in range(1,64,1):
            if (sheet.cell_value(i,1)==rno):
                for j in range(22):
                    print (sheet.cell_value(0, j)), "     :    ", (sheet.cell_value(i, j))
                for j in range(22, sheet.ncols):
                    sub = code_to_sub(str(sheet.cell_value(0, j)))
                    grade=str(sheet.cell_value(i, j))
                    mark, pt = grade_to_marks(grade)
                    print sub, "(", sheet.cell_value(0, j), ")", "      :      ", mark, "   ;   ",grade, "(", pt, ")"

def pfpercentage(sub):
        fail=0
        for i in range(22,sheet.ncols):
            if (sub==sheet.cell_value(0,i)):
                for j in range(1,63):
                    if (str(sheet.cell_value(j,i)) == "F" or str(sheet.cell_value(j,i)) == "FE"):
                        fail+=1
        failp=fail*100/63.0
        passp=100.0-failp
        print code_to_sub(sub),"(",sub,")"
        print "Pass Precentage :",passp,"%"
        print "Fail Precentage :",failp,"%"

def sub_wise_mrkl():
    c=0
    sub=str(raw_input("enter the subject name /subject code to print its mark list:"))
    subs=["CAL","CHE","EM","ICE","ISE","BE","CHE LAB","ICE WORKSHOP","BE WORKSHOP"]
    codes=["S1MA101","S1CY100","S1BE100","S1BE10105","S1BE103","S1EC100","S1CY110","S1CS110","S1EC110"]

    if (sub.isalpha()):    
        pos=subs.index(sub)
        for j in range (64):
            print  sheet.cell_value(c,0),"       ",sheet.cell_value(c,22+pos),"           ",sheet.cell_value(c,9)
            c=c+1

    else:
        pos=codes.index(sub)
        for j in range (64):
            print  sheet.cell_value(c,0),"       ",sheet.cell_value(c,22+pos),"           ",sheet.cell_value(c,9)
            c=c+1

    



print """

                                                                                        SCTCE
                                                                                Students Database
                                                                Computer Science First Year-R1
****************************************************************************************************************************************************
"""
while True :
    print """
                                                                                    Operations :
                                                                                1. Student details                     
                                                                                2. Subject wise mark list            
                                                                                3. Pass and Fail Percentage         
                                                                                4. Exit                                     
    """

    option=int(raw_input("Enter the option number:"))

    if (option==1):
        stud_details()
    elif (option==3):
        sub=str(raw_input("Enter the subject code or name:"))
        if sub not in ("S1MA101","S1CY100","S1BE100","S1BE10105","S1BE103","S1EC100","S1CY110","S1CS110","S1EC110","CAL","CHE","EM","ICE","ISE","BE","CHE LAB","ICE WRKSHOP","BE WRKSHOP"):
            print "Invalid subject"
        elif sub.isalpha():
            sub=sub_to_code(sub)
        pfpercentage(sub)
    elif (option==2):
        sub_wise_mrkl()    
    elif (option==4):
        break
    else:
        print "Invalid Option!"
