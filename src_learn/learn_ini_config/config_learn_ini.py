# основа работ с файлами config.ini
# Проблема в configparser что она принимает только текстовые значения
import configparser

config = configparser.ConfigParser()

config.add_section('Настройки')
config.set('Настройки', 'Ключ', 'True')
config.set('Excel', '')

with open('config.yaml', 'w') as file:




with open('config.ini', 'w', encoding="utf-8") as config_file:
    config.write(config_file)

user = config.get('Настройки', 'Ключ')
user
bool(user)