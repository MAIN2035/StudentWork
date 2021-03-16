import glob
import os
import zipfile
import time
import xlrd
import shutil
class CopyWork():
    def __init__(self):
        self.data = 0


    def import_excel(self,sourcePath):
        '''
            传入源文件路径，返回源数据
        :return:dataTable
        '''
        data = xlrd.open_workbook(sourcePath)
        data = data.sheets()[0]
        dataTable = []
        # print(self.data)

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


    def copy_file(self,nameDataTable,path1,path2):
        '''
        拷贝多个文件，并以指定名称命名
        :return:
        '''
        for i in nameDataTable:
            name = str(i['id'])+str(i['name'])+'.rar'
            path3 = path2 + name
            print(path3)
            shutil.copy(path1,path3)

    def zip(self,file,fileName,index):
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


if __name__ == '__main__':
    #存放命名信息的excle文件所在位置
    path='C:/Users/17630/Desktop/software_work/2019023578赵玉豪.xls'
    #将要复制的文件的位置
    path2='C:/Users/17630/Desktop/software_work/2019023578赵玉豪.rar'
    #存放复制文件的文件夹
    path3='C:/Users/17630/Desktop/software_work/work/'
    #压缩包的名称
    fileName='软件技术四班.zip'
    #压缩包存放的位置
    index='C:/Users/17630/Desktop/software_work/'

    # 获取命名信息
    nameDataTable = CopyWork().import_excel(path)
    #复制文件
    CopyWork().copy_file(nameDataTable,path2,path3)
    #压缩文件
    CopyWork().zip(path3,fileName,index)
    # for i in nameDataTable:
    # 1.拷贝并重命名文件
    # 2. 压缩并移动到目标目录


