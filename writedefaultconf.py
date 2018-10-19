from configparser import ConfigParser

CONFIG = ConfigParser()
CONFIG.read('Settings/config.ini')
CONFIG.add_section('email')
CONFIG.set('email', 'host', 'webmail.example.com')
CONFIG.set('email', 'user', 'Admin')
CONFIG.set('email', 'password', 'zabbix')
CONFIG.set('email', 'reports_folder', 'Inbox')
CONFIG.set('email', 'processed_folder', 'Processed')
CONFIG.add_section('elk')
CONFIG.set('elk', 'host', '192.168.0.1')
CONFIG.set('elk', 'port', '9200')

with open('Settings/config.ini', 'w') as f:
    CONFIG.write(f)
