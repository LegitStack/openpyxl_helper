# the purpose of this openpyxl helper module is to preserve charts from a template file.
# it will only preserve the charts on one sheet of your choosing.
# Here's how you use it:
# create a file like you normally would using openpyxl and save it as a
# different file than the one that has the charts on it you wish to preserve.
# this workaround will copy the charts from the original file and put it on the new file.
#
# invoke it by running this: (after you've already saved the file)
#   openpyxl_file.openpyxl_chart_saver(template_file_name, sheet_to_preserve, clean_name, new_name)
#
# argument considerations:
#   template_file_name indicates which file has charts in it you wish to transfer to the new file
#       template_file_name must be in the same folder as your script
#       template_file_name must the name of the file without the .xlsx at the end
#       template_file_name must be different than clean_name and new_name
#       default is "template"
#   sheet_to_preserve indicates which sheet you'd like to preserve charts on
#       sheet_to_preserve is the name of the sheet, its a string
#       default is "Sheet1"
#   clean_name indicates what name you'd like to use to clean the temp files this creates
#       clean_name is a string that will prepend every temp file this process creates
#       at the end of the process it will delete any file that starts with clean_name
#       so make sure its different than anything in your folder currently
#       for example if the template_file_name is "template" clean_name should not be "temp"
#       default is "_tempfile"
#   new_name indicates the new file you've created with openpyxl
#       new_name will be put in the same folder as your script
#       new_name must the name of the file without the .xlsx at the end (this filetype is assumed)
#       new_name must be different than clean_name and template_file_name
#       default is "finished_workbook"


import os
import shutil
import zipfile
import sys
import pandas as pd
from distutils.dir_util import copy_tree as imported_copy_tree
import time

def chart_saver(   template_file_name  =   "template"         ,
                            sheet_to_preserve   =   "Sheet1"           ,
                            clean_name          =   "_tempfile"        ,
                            new_name            =   "finished_workbook"):
    backup_charts(template_file_name, clean_name)
    restore_charts(sheet_to_preserve, new_name, clean_name)
    clean_up(clean_name)

def backup_charts(template_file_name, clean_name):
    # make a temporary template file, rename temp file as a zip
    shutil.copyfile('{0}.xlsx'.format(template_file_name), '{0}_temp-xlsx2.zip'.format(clean_name))

    # extract zip into temp directory so we can copy these files later
    zip_ref = zipfile.ZipFile('{0}_temp-xlsx2.zip'.format(clean_name), 'r')
    zip_ref.extractall("{0}_temp-zip".format(clean_name))
    zip_ref.close()

def restore_charts(sheet_to_preserve, new_name, clean_name):
    # take the file user created and turn it into a temp file
    os.rename("{0}.xlsx".format(new_name), '{0}_{1}'.format(clean_name, new_name))

    # make a zip out of the temp we created
    zip_name = '{0}_{1}_zip_file'.format(clean_name, new_name)
    zip_temp = '{0}_temp-zip'.format(clean_name)

    # copy and rename
    shutil.copyfile('{0}_{1}'.format(clean_name, new_name), "{0}_{1}.zip".format(clean_name, zip_name))

    # unzip new file
    zip_ref = zipfile.ZipFile("{0}_{1}.zip".format(clean_name, zip_name), 'r')
    zip_ref.extractall(zip_name)
    zip_ref.close()

    ### put data back:

    # 1. preserve drawings folder in xl
    if not os.path.exists(zip_name + "\\xl\\drawings"):
        os.makedirs(zip_name + "\\xl\\drawings")
    imported_copy_tree(zip_temp + "\\xl\\drawings", zip_name + "\\xl\\drawings")

    # 2. preserve charts folder in xl
    if not os.path.exists(zip_name + "\\xl\\charts"):
        os.makedirs(zip_name + "\\xl\\charts")
    imported_copy_tree(zip_temp + "\\xl\\charts", zip_name + "\\xl\\charts")

    # 3. preserve the folder: xl/worksheets/_rels
    if not os.path.exists(zip_name + "\\xl\\worksheets\\_rels"):
        os.makedirs(zip_name + "\\xl\\worksheets\\_rels")
    imported_copy_tree(zip_temp + "\\xl\\worksheets\\_rels", zip_name + "\\xl\\worksheets\\_rels")

    # 4. preserve file [content types].xml
    shutil.copy(zip_temp + "\\[Content_Types].xml", zip_name + "\\[Content_Types].xml") # has difficulty copying folders

    # 5. preserve the first worksheet file
    shutil.copy(zip_temp + "\\xl\\worksheets\\{0}.xml".format(sheet_to_preserve), zip_name + "\\xl\\worksheets\\{0}.xml".format(sheet_to_preserve)) # has difficulty copying folders

    # zip up
    shutil.make_archive("{0}_finished_{1}".format(clean_name, new_name), 'zip', zip_name)

    # rename it
    shutil.copyfile("{0}_finished_{1}.zip".format(clean_name, new_name), "{0}.xlsx".format(new_name))


def clean_up(clean_name):
    for each_file in os.listdir('.'):
        if each_file.startswith(clean_name):
            if os.path.isdir(each_file):
                shutil.rmtree(each_file)
            else:
                os.remove(each_file)


# can be called manually
if __name__ == '__main__':
    template_file_name  =  sys.argv[0]
    sheet_to_preserve   =  sys.argv[1]
    new_name            =  sys.argv[2]
    clean_name          =  sys.argv[3]

    backup_charts       (template_file_name, clean_name)
    restore_charts      (sheet_to_preserve, new_name, clean_name)
    clean_up            (clean_name)
