import openpyxl
import openpyxl_helper


# notice template_file_name and new_name are different!
def create_data_sheets(template_file_name="template", new_name="finished_workbook"):
    # choose a file as a template
    template_file = openpyxl.load_workbook('{0}.xlsx'.format(template_file_name))

    # define a data sheet
    sheet = template_file.get_sheet_by_name('Sheet2')

    # get some data
    stuff =[["MOSER, MARYANN", 376320.47, 65736.9113, 5608, 20118737],
            ["AIRMET, LORI", 371110.58, 77151.01036, 27058.48, 32752377],
            ["ALESSANDRI, MISTY", 347075.78, 53667.21784, 6316, 15307747.8],
            ["THURMAN, ANGIE", 318734.79, 71211.94821, 582.9, 17909119],
            ["JOHNSTON, CALOB", 297414.93, 53285.38886, 350, 13150670],
            ["LEWIS, TYLER", 202547.8, 36847.22583, 17200.86, 8400095],
            ["KIRKPATRICK, HALEY", 133876.48, 27161.91143, 319, 9855528],
            ["WESTBROOK, MICHAEL", 102652.65, 17408.90708, 33735.28, 5086840],
            ["MORRIS, BRIAN J", 99696.05, 13969.02274, 0, 3808744],
            ["HAWKINS, CHAD", 68057.17, 11880.20925, 0, 4340023]]

    # take this list and put it on the sheet object
    openpyxl_helper.list_of_lists_to_cell_range(list_of_lists=stuff,
                                                work_sheet=sheet,
                                                top_left='B4')

    # make a temp file with all the data in it.
    template_file.save('{0}.xlsx'.format(new_name))

create_data_sheets("template","finished_workbook")
openpyxl_helper.chart_saver(    template_file_name  =   "template"          ,
                                sheet_to_preserve   =   "Sheet1"            ,
                                clean_name          =   "_tempfile"         ,
                                new_name            =   "finished_workbook" )
