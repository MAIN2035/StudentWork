import glob
import os
import zipfile
import time
import xlrd
import shutil

def import_excel(sourcePath):
    '''
        传入源文件路径，返回源数据
    :return:dataTable
    '''
    data = xlrd.open_workbook(sourcePath)
    data = data.sheets()[0]
    dataTable = []
    # print(data)
    for rown in range(data.nrows):
        # 用字典存储学生信息，做短期存储
        array = {'flag': '', 'name': '', 'id': ''}
        #读取信息
        array['flag'] = data.cell_value(rown, 0)
        array['name'] = data.cell_value(rown, 2)
        array['id'] = data.cell_value(rown, 1)
        #每次装入一个学生数据到列表中,做长期封装
        dataTable.append(array)
    print(dataTable)
    return dataTable;
        # newname = id+name+format;
        # # print(newname)
        # #拷贝，并将文件以从excle中读取的数据命名
        # shutil.copy(oldname,newname)

def copy_file(nameDataTable,path1,path2):
    '''
    拷贝多个文件，并以指定名称命名
    :return:
    '''
    for i in nameDataTable:
        name = str(i['id'])+str(i['name'])+'.rar'
        path3 = path2 + name
        print(path3)
        shutil.copy(path1,path3)

def zip(file,fileName,index):
    '''
    以下路径请写绝对路径
    :param file:将要压缩的文件夹
    :param fileName: 压缩后压缩包的名称
    :param index: 压缩后的文件存放的位置
    :return:
    '''
    zipfile_name = os.path.basename(file) + fileName
    with zipfile.ZipFile(zipfile_name, 'w') as zfile:
        for foldername, subfolders, files in os.walk(file):
            fpath = foldername.replace(file, '')
            if file == foldername:
                for i in files:
                    zfile.write(os.path.join(foldername, i), os.path.join(fpath, i))
                    continue
                    zfile.write(foldername, fpath)
                    for i in files:
                        zfile.write(os.path.join(foldername, i), os.path.join(fpath, i))
        zfile.close()
    shutil.move(fileName,index)

def author():
    # 存放命名信息的excle文件所在位置
    path = 'C:/Users/17630/Desktop/software_work/2019023578赵玉豪.xls'
    # path = 'C:/Users/17630/Desktop/software_work/ymw.xls'
    # 将要复制的文件的位置
    path2 = 'C:/Users/17630/Desktop/software_work/2019023578赵玉豪.rar'
    # 存放复制文件的文件夹
    path3 = 'C:/Users/17630/Desktop/software_work/work/'
    # 获取命名信息
    nameDataTable = import_excel(path)
    # 复制文件
    copy_file(nameDataTable, path2, path3)
    # # 压缩文件
    # zip(path3, fileName, index)

def they():
    # 存放命名信息的excle文件所在位置
    path = input("请输入excle文件的位置：\t\t\t（回车键结束）").replace('\\', '/')
    # 将要复制的文件的位置
    path2 = input("请输入源文件的位置：\t\t\t（回车键结束）").replace('\\', '/')
    # 存放复制文件的文件夹
    path3 = input("复制后的文件将放在哪里？\t\t\t（回车键结束）").replace('\\', '/')
    # 获取命名信息
    nameDataTable = import_excel(path)
    #复制文件
    copy_file(nameDataTable, path2, path3)


if __name__ == '__main__':
    flag = int(input("是否为默认配置？\t\t\t（是，填1/否，填0）"))
    if(flag):
        author()
    else:
        they()







