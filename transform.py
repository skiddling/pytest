import xlrd
import xlwt

def transform():
    in_file = xlrd.open_workbook('data.xls')
    out_file = xlwt.Workbook()
    in_sheet = in_file.sheets()[0]
    out_sheet = out_file.add_sheet('data')

    in_rows = in_sheet.nrows
    in_cols = in_sheet.ncols

    teacher_course_dic = {}

    for i in range(1, in_rows):
        for j in range(3, in_cols, 2):
            if (in_sheet.cell_type(i, j) not in (xlrd.XL_CELL_BLANK,
                                                       xlrd.XL_CELL_EMPTY)):
                course = in_sheet.cell(i, j).value
                teacher = in_sheet.cell(i, j + 1).value
                teacher_course_dic[teacher] = course

    #k = 0
    #for teacher in teacher_course_dic:
    #    #print teacher, teacher_course_dic
    #    out_sheet.write(k, 0, teacher)
    #    out_sheet.write(k, 1, teacher_course_dic[teacher])
    #    k += 1

    dic = {}
    for i in range(1, in_rows):
        class_name = in_sheet.cell(i, 1).value
        class_time = int(in_sheet.cell(i, 2).value)
        day_time = 1
        for j in range(4, in_cols, 2):
            if (in_sheet.cell_type(i, j) not in (xlrd.XL_CELL_BLANK,
                                                 xlrd.XL_CELL_EMPTY)):
                teacher = in_sheet.cell(i, j).value
                if(dic.has_key(class_name) == False):
                    dic[class_name] = {}
                if(dic[class_name].has_key(teacher) == False):
                    dic[class_name][teacher] = []
                dic[class_name][teacher].append(str(day_time) + '-' + str(class_time))
            day_time += 1

    k = 0
    for cln in dic:
        for ten in dic[cln]:
            t = 2
            out_sheet.write(k, 0, cln + teacher_course_dic[ten])
            out_sheet.write(k, 1, ten)
            s = ''
            for case in dic[cln][ten]:
                if(s == ''):
                    s += case
                else :
                    s += ',' + case

                #out_sheet.write(k, t, case)
                #t += 1
            out_sheet.write(k, 2, s)
            out_sheet.write(k, 3, cln[0:2] + teacher_course_dic[ten])
            out_sheet.write(k, 4, cln)
            k += 1
    out_file.save('out_data.xls')

transform()
