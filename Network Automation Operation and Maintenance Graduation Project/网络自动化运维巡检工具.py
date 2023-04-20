import pandas as pd                            # 引入pandas模块对EXCEL数据进行处理
from netmiko import ConnectHandler, redispatch  # 引入netmiko模块，实现远程连接设备并下发命令
from fuzzywuzzy import fuzz                    # 对数据进行模糊字符串匹配
from pathlib import Path as path               # 引入pathlib库中的Path模块，实现文件以及目录的操作
from datetime import datetime                  # 获取日期时间
from multiprocessing.pool import ThreadPool    # 实现多线程（在测试时存在BUG，未使用）


class backupconfig(object):
    def __init__(self, file_name):
        """
        定义使用的EXCEL名
        定义线程池中的线程数
        """
        self.file_name = file_name
        # self.pool = ThreadPool(10)

    def get_device_info(self):  # 信息模块
        """
        获取设备信息
        返回一个生成器
        """
        try:
            wb = pd.read_excel(self.file_name, "设备信息",
                               usecols="A:H", keep_default_na=False)
            data = wb.to_dict(orient="records")
            for row in data:
                if fuzz.partial_ratio("SW", row['name']) == 100:
                    row['cmd_list'] = self.get_sw_cmd_info()
                elif fuzz.partial_ratio("FW", row['name']) == 100:
                    row['cmd_list'] = self.get_fw_cmd_info()
                else:
                    pass
                yield row

        except FileNotFoundError:
            print(f"{self.file_name} 文件不存在！")

    def get_sw_cmd_info(self):  # 信息模块
        """
        获取交换机巡检命令信息，返回一个列表，其中包含交换机上需要执行的命令
        """
        try:
            sw_cmd_list = []
            cb = pd.read_excel(self.file_name, 'huawei_巡检',
                               usecols="A:B", keep_default_na=False)
            data = cb.to_dict(orient="records")
            for cmd_row in data:
                if ("all" in cmd_row['Comment'] or "SW" in cmd_row['Comment']):
                    sw_cmd_list.append(cmd_row['Command'].strip())

            return sw_cmd_list

        except FileNotFoundError:
            print(f"{self.file_name} 文件不存在！")

    def get_fw_cmd_info(self):  # 信息模块
        """
        获取防火墙巡检命令信息,返回一个列表，其中包含防火墙上需要执行的命令
        """
        fw_cmd_list = []
        cb = pd.read_excel(self.file_name, 'huawei_巡检',
                           usecols="A:B", keep_default_na=False)
        data = cb.to_dict(orient="records")
        for cmd_row in data:
            if ("all" in cmd_row['Comment'] or "FW" in cmd_row['Comment']):
                fw_cmd_list.append(cmd_row['Command'].strip())

        return fw_cmd_list

    def connect_jump(self):  # 连接模块
        """
        本地连接堡垒机
        本地创建一个存放日志的文件夹，文件地址为“./log/{Current date}”
        连接成功后，会显示登录成功
        返回一个长连接对象
        """
        for host in self.get_device_info():
            if host['Comment'] == '#':
                del host
                continue
            if host['Comment'] == '堡垒机':
                if host['protocol'].lower().strip() == 'ssh':
                    host['port'] = host['port'] if (
                        host['port'] != 22 and host['port']) else 22
                    host.pop('name'),
                    host.pop('Comment'),

                if host['protocol'].lower().strip() == 'telnet':
                    host['port'] = host['port'] if (
                        host['port'] != 23 and host['port']) else 23
                    host.pop('name'),
                    host.pop('protocol'),
                    host.pop('Comment')

                host.pop('protocol')
                if 'huawei' in host['device_type']:
                    connect = ConnectHandler(**host, conn_timeout=15)
                break
        path(
            f"./log/{datetime.now().strftime('%Y-%m-%d')}").mkdir(parents=True, exist_ok=True)
        return connect

    def connect_device(self, host):  # 应用模块、输出模块
        """
        接收host参数，其中包括需要巡检的设备信息以及命令列表
        通过在堡垒机上使用telnet连接需要巡检的设备，下发命令，并将回显的命令保存在本地文件中
        保存每台设备的配置文件到“./config_backup”
        """
        if host['Comment'] in ['#', '堡垒机']:
            return 0
        if host["protocol"].lower().strip() == "telnet":
            host['port'] = host['port'] if (
                host['port'] != 23 and host['port']) else 23

            connect = self.connect_jump()
            print(connect.find_prompt())

            connect.username = host['username']
            connect.password = host['password']
            connect.write_channel(f"telnet {host['ip']} {host['port']} \n")
            connect.telnet_login()
            print(
                f"\n----------------- Successfully logged in to {host['ip']} -----------------")
            print(connect.find_prompt())
            redispatch(connect, device_type=host['device_type'])
            path(f"./log/{datetime.now().strftime('%Y-%m-%d')}/{host['name']}").mkdir(
                parents=True, exist_ok=True)
            for cmd in host["cmd_list"]:
                output = connect.send_command(
                    cmd, strip_prompt=False, strip_command=False)
                path(
                    f"./log/{datetime.now().strftime('%Y-%m-%d')}/{host['name']}/{cmd}.txt").write_text(output)
                # print(output)
            config_backup = connect.send_command("dis cu")
            path(
                f"./config_backup/{host['name']}_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt").write_text(config_backup)
            connect.disconnect()

    def run_cmd(self):  # 应用模块
        """
        获取设备信息，并将长连接中的对象发送给connect_device
        使用多线程减少程序运行时间（测试时存在BUG，未使用）
        """
        hosts = self.get_device_info()
        for host in hosts:
            self.connect_device(host)
        #     self.pool.apply_async(self.connect_device,args=[host,])
        # self.pool.close()
        # self.pool.join()


if __name__ == '__main__':
    path("./log").mkdir(parents=True, exist_ok=True)
    path("./config_backup").mkdir(parents=True, exist_ok=True)
    file_name = input("Please enter your file name or absolute path:")
    net = backupconfig(file_name=file_name)
    net.run_cmd()
