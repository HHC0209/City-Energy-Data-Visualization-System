from influxdb import InfluxDBClient
import win32api
import os

win32api.ShellExecute(0, 'open', 'influxdb-1.8.2-1\\influxd.exe', '', '', 0)  # 运行influxd.exe
client = InfluxDBClient('localhost', 8086, 'root', 'root', 'Energy')

client.query('drop measurement "Data"')
for root, dirs, files in os.walk(r'configs'):
    for file in files:
        os.remove(os.path.join(root, file))
for root, dirs, files in os.walk(r'funcs'):
    for file in files:
        os.remove(os.path.join(root, file))
for root, dirs, files in os.walk(r'graphs'):
    for file in files:
        os.remove(os.path.join(root, file))





