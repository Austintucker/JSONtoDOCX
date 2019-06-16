# Adapted From https://stackoverflow.com/questions/48150222/changing-paragraph-formatting-in-python-docx

from docx import Document
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.style import WD_STYLE_TYPE
import json, os

def resumeDocxConstructor():

    alignment_dict = {  'justify': WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
                        'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
                        'right': WD_PARAGRAPH_ALIGNMENT.RIGHT,
                        'left': WD_PARAGRAPH_ALIGNMENT.LEFT}

    document = Document()

    sections = document.sections
    for section in sections:
        section.top_margin      = Inches(1)
        section.bottom_margin   = Inches(1)
        section.left_margin     = Inches(1)
        section.right_margin    = Inches(1)

    ####  Temp  #### - Eventually this information will be passed by HTTP trigger

    resumeDirectory = f"{os.getcwd()}/JSON"

    addressDoc  = f"{resumeDirectory}/address.json"
    resumeDoc   = f"{resumeDirectory}/resume.json"
    userDoc     = f"{resumeDirectory}/user.json"

    with open(addressDoc, "r") as read_file:
        address = json.load(read_file)
    with open(resumeDoc, "r") as read_file:
        resume = json.load(read_file)
    with open(userDoc, "r") as read_file:
        user = json.load(read_file)

    resumeOrder = [1,2,3,4,5]   # This list will determine the order in which sections are added to the resume
                                # should eventually get this from the JSON

    includeAddress = True # This param will be given in http trigger

    font ='Calibri'

    #### *Temp* ####


    #### Define Functions ####

    # add_content function adapted From https://stackoverflow.com/questions/48150222/changing-paragraph-formatting-in-python-docx
    def add_content(content, space_after, font_name=font, font_size=16, line_spacing=0, space_before=0,
                    align='left', keep_together=True, keep_with_next=False, page_break_before=False,
                    widow_control=False, set_bold=False, set_italic=False, set_underline=False, set_all_caps=False, 
                    style_name="", firstline_indent=0.0, left_indent=0.0):
        paragraph = document.add_paragraph(content)
        paragraph.style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
        font = paragraph.style.font
        font.name = font_name
        font.size = Pt(font_size)
        font.bold = set_bold
        font.italic = set_italic
        font.all_caps = set_all_caps
        font.underline = set_underline
        paragraph_format = paragraph.paragraph_format
        paragraph_format.alignment = alignment_dict.get(align.lower())
        paragraph_format.space_before = Pt(space_before)
        paragraph_format.space_after = Pt(space_after)
        paragraph_format.line_spacing = line_spacing
        paragraph_format.keep_together = keep_together
        paragraph_format.keep_with_next = keep_with_next
        paragraph_format.page_break_before = page_break_before
        paragraph_format.widow_control = widow_control
        paragraph_format.first_line_indent = Inches(firstline_indent)
        paragraph_format.left_indent = Inches(left_indent)


    def generateHeader(): # Name, Address, Contact Info, *perhaps should add websites here*
        # Insert Name
        add_content(f"{user['FName']} {user['LName']}",
                    align='Center', space_before=0, space_after=0, line_spacing=1, font_size=16, set_bold=True, set_all_caps=True,style_name ="NameBold", firstline_indent=0.0, left_indent=0.0)

        # Insert Address
        if includeAddress:
            if address['Address']['Country'] != 'USA':
                country = address['Address']['Country']
            else:
                country = ""

            add_content(f"{address['Address']['Line1']} {address['Address']['Line2']}\n"
                        f"{address['Address']['City']}, {address['Address']['State']} {address['Address']['Zip']} {country}",
                        align='Center', space_before=0, space_after=0, line_spacing=1, font_size=12, set_bold=False, set_all_caps=False,style_name ="addressNotBold", firstline_indent=0.0, left_indent=0.0)

        # Insert Contact Info
        add_content(f"Phone: {user['Phone']}\n"
                    f"Email: {user['Email']}",
                    align='Center', space_before=0, space_after=0, line_spacing=1, font_size=12, set_bold=False, set_all_caps=False,style_name ="contactNotBold", firstline_indent=0.0, left_indent=0.0)


    def generateObjective():
        # Insert Objective Statement
        add_content(f"\nObjective:",
                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=14, set_bold=True, set_all_caps=False,style_name ="ObjectiveBold", firstline_indent=0.0, left_indent=0.0)
        add_content(f"{resume['Objective Statement']}\n",
                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=12, set_bold=False, set_all_caps=False,style_name ="ObjectiveStatement", firstline_indent=0.0, left_indent=0.5)


    def generateEducation():
        # Insert Education
        add_content(f"Education:",
                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=14, set_bold=True, set_all_caps=False,style_name ="EducationBold", firstline_indent=0.0, left_indent=0.0)

        j = 0
        for i in resume['School']:
            add_content(f"\t{i['Name']} - {i['City']}, {i['State']}, {i['Country']}\n"
                        f"\tMajor: {i['Major']}\t\tGraduation: {i['Graduation']}\n"
                        f"\tMinor: {i['Minor']}\n"
                        f"\tGPA: {i['Gpa']}\n",
                        align='Left', space_before=0, space_after=0, line_spacing=1, font_size=12, set_bold=False, set_all_caps=False,style_name =f"ObjectiveStatement{j}", firstline_indent=0.0, left_indent=0.0)
            j += 1


    def generateCoursework():
        # Insert Relevant Coursework
        add_content(f"Relevant Coursework:",
                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=14, set_bold=True, set_all_caps=False,style_name ="CourseworkBold", firstline_indent=0.0, left_indent=0.0)

        j = 0
        for i in resume['RelevantCourse']:
            add_content(f"\t+ {i['Name']}",
                        align='Left', space_before=0, space_after=0, line_spacing=1, font_size=12, set_bold=False, set_all_caps=False,style_name =f"CourseName{j}", firstline_indent=0.0, left_indent=0.0)
            add_content(f"{i['Description']}",
                        align='Left', space_before=0, space_after=0, line_spacing=1, font_size=10, set_bold=False, set_all_caps=False,style_name =f"CourseDes{j}", firstline_indent=0.0, left_indent=1.0)
            j += 1

        add_content(f"",
                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=14, set_bold=True, set_all_caps=False,style_name ="CourseworkBlankLine", firstline_indent=0.0, left_indent=0.0)


    def generateSkills(): # Should experience level be included or is that more for analysis
        # Insert Skills - This section is just bad and needs work, might be worth looking at tables
        add_content(f"Skills:",
                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=14, set_bold=True, set_all_caps=False,style_name ="SkillsBold", firstline_indent=0.0, left_indent=0.0)

        for i in resume['Skill']:
            add_content(f"\t{i}",
                        align='Left', space_before=0, space_after=0, line_spacing=1, font_size=12, set_bold=False, set_all_caps=False,style_name =f"{i}", firstline_indent=0.0, left_indent=0.0)
            for j in resume['Skill'][i]:
                for k in resume['Skill'][i][j]:
                        add_content(f"\t\t{k} - {resume['Skill'][i][j][k]}",
                                    align='Left', space_before=0, space_after=0, line_spacing=1, font_size=10, set_bold=False, set_all_caps=False,style_name =f"{k}", firstline_indent=0.0, left_indent=0.0)


    def generateExperience():
        return


    def generateActivities():
        return

    functionDic = { 1: generateHeader,
                    2: generateObjective,
                    3: generateEducation,
                    4: generateCoursework,
                    5: generateSkills}

    for x in resumeOrder:
        functionDic[x]()


    # Save file to a /local directory
    os.chdir(f"{os.getcwd()}/resumeDoc")
    document.save(f"{user['LName']}.docx")


### Execute ###

resumeDocxConstructor()
#Test Changed
# Ian Change 4