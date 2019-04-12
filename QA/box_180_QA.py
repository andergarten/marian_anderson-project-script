import os
import shutil
import subprocess
from openpyxl import *


def refold_and_rename(path):
    '''This function is for refold_and_rename
    for box 180
    '''

    # iterate over file list in folder
    for file_name in os.listdir(path):

        # pass CaptureOne folder and .DS_Store file
        if file_name != "CaptureOne" \
           and file_name != ".DS_Store" \
           and "_" in file_name: #???

            # obtain ark id
            ark_id = ''
            for charactor in file_name:
                if charactor != '_':
                    ark_id += charactor
                else:
                    break

            # make new folder with ark id
            new_path = path + "\\" + ark_id
            if not os.path.exists(new_path):
                os.makedirs(new_path)

            # move files and rename
            old_file = path + "\\" + file_name
            
            # if it is not a sample image, rename
            if "testref1" not in file_name:
                
                # if it is a regular page, keep last 8 digits
                if 'a.' not in file_name:
                    new_file = new_path + "\\" + file_name[-8:]

                # if it is named irregularly, keep last 9 digits
                else:
                    new_file = new_path + "\\" + file_name[-9:]
                    
            # otherwise, do not rename
            else:
                new_file = new_path + "\\" + file_name

            # move to new folder named as ark_id    
            shutil.move(old_file, new_file)

'''
def get_page_number(path, ark_id, page_number):

    if page_number == 0:
        
        # images path inside ark_id folder
        image_path = path + ark_id
         
        # calculate how many page in the folder (exclude sample image)
        page_number = len(os.listdir(image_path)) - 1

        # if it is named irregularly, page -1, we do not want that page
        for image_name in os.listdir(image_path):
            if 'a.' in file_name:
                page_number -= 1

    return page_number
'''

def create_spreadsheet(path, ark_id, page_number):

    # images path inside ark_id folder
    image_path = path + ark_id

    # copy spreadsheet template file to folder path
    shutil.copyfile(r"T:\MarianAnderson\Copy of Marian Anderson spreadsheet template.xlsx",\
                    path + "\\Copy of Marian Anderson spreadsheet template.xlsx")

    # change cwd to folder_path
    os.chdir(path)
        
    # open template
    template = load_workbook\
               ('Copy of Marian Anderson spreadsheet template.xlsx')

    # get the first worksheet
    sheet1 = template["Structural Metadata"]

    '''
    # get page number for each ark_id item
    page_number = get_page_number(box_number, folder_number, ark_id, page_number)
    '''
    
    # use loop to fill sheet
    for page in range(1, page_number+1):

        # fill "visible page" field
        sheet1['B'+ str(3 + page)] = page;

        # fill "filename" field
        sheet1['D'+ str(3 + page)] = '000' + str(page) + '.tif'

    # save updated file
    template.save('81431' + ark_id + '.xlsx')

def refine_arklist(path):
    
    # get existing xlsx file in a list
    xlsx_list = [x for x in os.listdir(path) \
                if os.path.isfile(os.path.join(path,x))]

    # loop over the ark_id folders
    ark_list = [a for a in os.listdir(path) \
                  if not os.path.isfile(os.path.join(path,a))]

    #update ark_id_list
    for xlsx_name in xlsx_list:
        for ark_id in ark_list:
            #if spreadsheet has already been created, skip it
            if ark_id in xlsx_name:
                #print (xlsx_name, ark_id)
                ark_list.remove(ark_id)

    return ark_list

def qa(path):

    ark_list = refine_arklist(path)
   
    for ark_id in ark_list:

        # exclude non-ark_id folder
        if "CaptureOne" not in ark_id:

            # ask student to do QA
            qa_result = input("Does " + ark_id + \
                              " pass QA? y for pass, n for fail.\n")

            while (qa_result != "y" and qa_result != "n"):
                print (qa_result)

                qa_result = input("Wrong input. y for pass, n for fail.\n")
                
            # if pass, create spreadsheet
            if qa_result == "y":
                # tell student what to do
                print("Please put ‘Y’ into ‘QA Pass’ field.\n")
                # ask for page number
                page_number = int(input("What is the page number?\n"))
                # create a spreadsheet
                create_spreadsheet(path, ark_id, page_number)

            # if fail, 
            elif qa_result == "n":
                # tell student what to do
                print("Please put ‘N’ into ‘QA Fail’ field and fill in reasons. Email Craig with description of problem.\n")
                # ask for page number
                page_number = int(input("What is the correct page number?\n"))
                # create a spreadsheet
                create_spreadsheet(path, ark_id, page_number)

    # remove extra file
    copy_path = path + "\\Copy of Marian Anderson spreadsheet template.xlsx"
    if os.path.exists(copy_path):
        os.remove(copy_path)


def main():
    
    # install openpyxl package
    subprocess.check_call(["python", '-m', 'pip', 'install', 'openpyxl'])

    # ask for drive number
    drive_letter = input('What is the drive letter for sceti-completed?\n')

    # ask for box number
    box_number = str(input('Give me a box number.\n'))

    while True:

        # get folder number
        folder_number = str(input('Give me a 4-digit folder number.\n'))
        
        # target folder
        path = drive_letter + r":\MarianAnderson\mscoll200_box" + box_number \
               + "\\folder0" + folder_number
        
        # refold and rename
        refold_and_rename(path)

        print("Refolding and renaming are done...\n")

        # do QA
        qa(path)

        print("QA and creating spreadsheets are done for folder" + \
              folder_number + "...\n")
        # tell student to upload spreadsheet
        print("Now please upload spreadsheets in folder 0" + \
              folder_number + " to Google Drive...\n")

        # get box number
        box_number = str(input('Give me a new box number.\n'))
                    
if __name__ == '__main__':
    main()
