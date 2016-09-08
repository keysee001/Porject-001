import os
import io
import excel
import xlrd
import xlwt
import re
import os.path




def record(string, filename):
    f = open(filename, "a")
    # print os.path._getfullpathname(f.name)
    f.write(string + '\n' + '\n')
    f.close()


def clear(filename):
    f = open(filename, "w")
    f.write('')
    f.close()


# count lines in a file
class FileCount(object):
    report = []
    finalReport =[]
    fileType = ['.java', '.py', '.html', '.js']
    def countLines(self, rootDir):
        f = None
        lines = []
        javaCodeLines = 0
        pythonCodeLines = 0
        jsCodeLines = 0
        htmlCodeLines = 0

        file_type = None
        list_files = os.listdir(rootDir)

        for root, dirs, files in os.walk(rootDir):
            for child in files:
               path = os.path.join(root, child)
               if os.path.splitext(path)[1] in FileCount.fileType:
                    file_type = os.path.splitext(path)[1].split('.')[1]
                    print file_type
                    if file_type == 'java':
                        (whiteLines, commentLine, normalLines) = java_codeCount(str(path))
                    elif file_type == 'py':
                        (whiteLines,commentLine,normalLines) = py_codeCount(str(path))
                    elif file_type == 'html':
                        (whiteLines, commentLine, normalLines) = html_codeCount(str(path))
                    else:
                        (whiteLines, commentLine, normalLines) = js_codeCount(str(path))

                    FileCount.report.append([file_type, child, str(whiteLines),str(commentLine),str(normalLines), path])


                    print ' FilePath : ' + str(path) + '\n' + ' File type : ' + str(
                        file_type) + '\n' + ' FileName : ' + str(child) + '\n' + ' Total white lines : ' + str(
                        whiteLines)+'\n'+'Total comment lines :' + str(commentLine)+'\n'+'Total code lines :'+ str(normalLines)+'\n'



        for subreport in FileCount.report:
            if subreport[0]== 'java':
                javaCodeLines = javaCodeLines + int(subreport[4])
            elif subreport[0]== 'py':
                pythonCodeLines = pythonCodeLines +  int(subreport[4])
            elif subreport[0]== 'js':
                jsCodeLines = jsCodeLines + int(subreport[4])
            elif subreport[0]== 'html':
                htmlCodeLines += int(subreport[4])

        FileCount.finalReport=[['javaCodeLines',javaCodeLines],['pythonCodeLines',pythonCodeLines],['jsCodeLines',jsCodeLines],['htmlCodeLines',htmlCodeLines]]



        return FileCount.report,FileCount.finalReport

def py_codeCount(filename):
    whiteLines = 0
    commentLine = 0
    normalLines = 0
    f= open(filename,'r')
    comment = False
    for line in f.readlines():
        line = line.strip()
        if not line.split():
            whiteLines = whiteLines + 1
        elif line.startswith("'''") or line.startswith('"""') and  not line.endswith("'''") and not line.endswith('"""'):
            commentLine = commentLine +1
            comment = True
        elif comment == True :
            commentLine = commentLine + 1
            if line.endswith("'''") or line.endswith('"""'):
                comment = False
        elif line.startswith('#'):
            commentLine = commentLine + 1
        else :
            normalLines = normalLines + 1
    return (whiteLines,commentLine,normalLines)



def java_codeCount(filename):
    whiteLines = 0
    commentLine = 0
    normalLines = 0
    f= open(filename,'r')
    comment = False
    for line in f.readlines():
        line = line.strip()
        if not line.split():
            whiteLines = whiteLines + 1
        elif line.startswith("/*") or line.startswith('/**') and  not line.endswith("*/"):
            commentLine = commentLine +1
            comment = True
        elif comment == True :
            commentLine = commentLine + 1
            if line.endswith("*/"):
                comment = False
        elif line.startswith('//'):
            commentLine = commentLine + 1
        else :
            normalLines = normalLines + 1
    return (whiteLines,commentLine,normalLines)


def js_codeCount(filename):
    whiteLines = 0
    commentLine = 0
    normalLines = 0
    f= open(filename,'r')
    comment = False
    for line in f.readlines():
        line = line.strip()
        if not line.split():
            whiteLines = whiteLines + 1
        elif line.startswith("/*") or line.startswith('/**') and  not line.endswith("*/"):
            commentLine = commentLine +1
            comment = True
        elif comment == True :
            commentLine = commentLine + 1
            if line.endswith("*/"):
                comment = False
        elif line.startswith('//') or line.startswith('<!-'):
            commentLine = commentLine + 1
        else :
            normalLines = normalLines + 1
    return (whiteLines,commentLine,normalLines)

def html_codeCount(filename):
    whiteLines = 0
    commentLine = 0
    normalLines = 0
    f= open(filename,'r')
    comment = False
    for line in f.readlines():
        line = line.strip()
        if not line.split():
            whiteLines = whiteLines + 1
        elif line.startswith("<!--") and  not line.endswith("-->"):
            commentLine = commentLine +1
            comment = True
        elif comment == True :
            commentLine = commentLine + 1
            if line.endswith("-->"):
                comment = False
        elif line.startswith("<!--")and line.endswith("-->"):
            commentLine = commentLine + 1
        else :
            normalLines = normalLines + 1
    return (whiteLines,commentLine,normalLines)


def create_excel(report,finalReport):
    file = xlwt.Workbook()
    table = file.add_sheet('eachFileCodeLineReport', cell_overwrite_ok=True)
    headerStyle=xlwt.easyxf('font:height 240, color-index black, bold on;align: wrap on, vert centre, horiz center')
    #bodyStyle=xlwt.easyxf('font:height 100, color-index black,align: wrap on, vert centre, horiz center')
    #headerStyle = file.add_format({'border': 1, 'align': 'center', 'bg_color': 'cccccc', 'font_size': 13, 'bold': True})
    table.write(0, 0, "FileType", headerStyle);
    table.write(0, 1, "FileName", headerStyle);
    table.write(0, 2, "WhiteSpaceNumber", headerStyle);
    table.write(0, 3, "CommentNumber", headerStyle);
    table.write(0, 4, "NormalCodeNumber", headerStyle);
    table.write(0, 5, "FilePath", headerStyle);
    recordRows=len(report)
    recordRow=1
    #while recordRow >recordRows:
    for sublist in report:
         if recordRow<= recordRows:
            for recordCol in range(len(sublist)):
                table.write(recordRow,recordCol,sublist[recordCol])
            recordRow=recordRow+1




    table1 = file.add_sheet('eachFileTypeCodeLineReport', cell_overwrite_ok=True)
    headerStyle=xlwt.easyxf('font:height 240, color-index black, bold on;align: wrap on, vert centre, horiz center')
    table.write(0, 0, "FileType", headerStyle);
    table.write(0, 1, "FileName", headerStyle);
    table.write(0, 2, "WhiteSpaceNumber", headerStyle);
    table.write(0, 3, "CommentNumber", headerStyle);
    table.write(0, 4, "NormalCodeNumber", headerStyle);
    table.write(0, 5, "FilePath", headerStyle);
    recordRows=len(finalReport)
    recordRow=1
    for sublist in finalReport:
        if recordRow<= recordRows:
             for recordCol in range(len(sublist)):
                 table1.write(recordRow,recordCol,sublist[recordCol])
             recordRow=recordRow+1

    file.save('Report.xls')














if "__main__" == __name__:
    # clear('codelineCount.txt')
    if os.path.exists('Report.xls'):
        os.remove('C:\Python_Test\util\Report.xls')

    a = FileCount()
    a.countLines('C:\Python_Test\util')
    print a.report
    print a.finalReport
    create_excel(a.report,a.finalReport)
