# python
# -*- coding:utf-8 -*-
# Author:源~
# 2019年11月05日
# 逻辑思路：调用vimc的Windows管理技术命令接口，获取相关信息
# 2019年11月05日，初步逻辑搭建完成，基本命令测试成功
# 目前获取的有：操作系统信息、版本、位数、IP地址、
# 2019年11月07日，重构相关代码，同时增加数据库写入功能，分为硬件表和软件表
# 2020年08月20日，连接数据库成功，运行成功
import getpass
import math
import sys

import wmi
import socket
import platform
import winreg
import pymongo


class Get_Pc_Date(object):
    def __init__(self):
        self.c = wmi.WMI()

    # 系统信息
    def get_system_date(self):
        system_name = platform.platform()[:-(len(platform.version()) + 1)]  # 操作系统
        system_version = platform.version()  # 操作系统版本号
        system_digit = platform.architecture()[0]  # 操作系统位数
        return system_name, system_version, system_digit

    # CPU信息
    def get_CPU(self):
        for cpu in self.c.Win32_Processor():
            cpu_number = cpu.NumberOfLogicalProcessors  # 逻辑处理器数量
            cpu_name = cpu.Name  # 处理器型号
        return cpu_number, cpu_name

    # 内存信息
    def get_PhysicalMemory(self):
        for mem in self.c.Win32_PhysicalMemory():
            ram_number = mem.SerialNumber  # 内存编号
            ram_size = str(math.ceil(int(mem.Capacity) / 1024 ** 3)) + "GB"  # 内存大小，四舍五入计算
        return ram_number, ram_size

    # BIOS信息
    def get_video(self):
        for v in self.c.Win32_BIOS():
            bios_number = v.SerialNumber  # BIOS出厂编号
            bios_name = v.Name  # BIOS名称
        return bios_number, bios_name

    # 获得主板信息
    def get_BaseBoard(self):
        for i in self.c.Win32_BaseBoard():
            main_board_model = i.Product  # 主板产品型号
            main_board_company = i.Manufacturer  # 主板厂家名字
            main_board_name = i.Name  # 主板名称
        return main_board_model, main_board_company, main_board_name

    # 硬盘
    def printDisk(self):
        for disk in self.c.Win32_DiskDrive():
            disk_size = str(math.ceil(int(disk.Size) / 1024 ** 3)) + "GB"  # 磁盘大小
            disk_company = disk.Model  # 磁盘驱动器的制造商的型号
        return disk_size, disk_company

    # 获得电脑安装软件信息
    def get_software_data(self):
        # 需要遍历的两个注册表,此处目的是为了避免系统不是64位,即为兼容性而设计
        sub_key = [r'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall',
                   r'SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall']

        software_name = []

        for i in sub_key:
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, i, 0, winreg.KEY_ALL_ACCESS)
            for j in range(0, winreg.QueryInfoKey(key)[0] - 1):
                try:
                    key_name = winreg.EnumKey(key, j)
                    key_path = i + '\\' + key_name
                    each_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path, 0, winreg.KEY_ALL_ACCESS)
                    DisplayName, REG_SZ = winreg.QueryValueEx(each_key, 'DisplayName')
                    DisplayName = DisplayName.encode('utf-8')
                    software_name.append(DisplayName)
                except WindowsError:
                    pass

        # 去重排序
        software_name = list(set(software_name))  # 消除重复元素
        software_name = sorted(software_name)  # 排序

        a = 0
        sortware_name = []
        for result in software_name:
            a += 1
            result = result.decode()  # 转换编码
            sortware_name.append(result)
        # a为最终的软件数量
        return sortware_name, a

    # 将得到的数据转换成字典格式
    def get_dic(self):
        pc_dic = {}
        system = self.get_system_date()
        pc_dic['操作系统'] = system[0]
        pc_dic['操作系统版本号'] = system[1]
        pc_dic['操作系统位数'] = system[2]
        cpu = self.get_CPU()
        pc_dic['逻辑处理器数量'] = cpu[0]
        pc_dic['处理器型号'] = cpu[1]
        physicalmemor = self.get_PhysicalMemory()
        pc_dic['内存编号'] = physicalmemor[0]
        pc_dic['内存大小'] = physicalmemor[1]
        video = self.get_video()
        pc_dic['BIOS出厂编号'] = video[0]
        pc_dic['BIOS名称'] = video[1]
        printdisk = self.printDisk()
        pc_dic['磁盘大小'] = printdisk[0]
        pc_dic['磁盘制造商型号'] = printdisk[1]
        baseboard = self.get_BaseBoard()
        pc_dic['主板产品型号'] = baseboard[0]
        pc_dic['主板厂家名字'] = baseboard[1]
        pc_dic['主板名称'] = baseboard[2]
        return pc_dic

    # 主逻辑函数
    def main(self):
        try:
            user_name = getpass.getuser()
            hostname = socket.getfqdn(socket.gethostname())
            ip = socket.gethostbyname(hostname)  # IP地址
            pc_dic = self.get_dic()  # 获取硬件信息
            pc_sortware = self.get_software_data()  # 获得软件信息
            pc = {'user_name': user_name, 'network': ip[0:8], 'ip': ip, 'computer_hardware': pc_dic,
                  'software_name': pc_sortware[0],
                  'software_number': pc_sortware[1]}
            client = pymongo.MongoClient(host='10.85.7.140', port=27017)
            db = client.cd_pc_data  # 这是数据表的名字
            collection = db.pc_data  # 这是集合
            collection.insert_one(pc)
        except:
            sys.exit()


if __name__ == '__main__':
    get_pc_date = Get_Pc_Date()
    get_pc_date.main()
