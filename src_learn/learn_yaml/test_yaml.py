import yaml


# Замена данных вариант 2
name_config_file = 'test_config.yaml'
key = 'key5'
new_value = 'new_value5'

def set_config(name_config_file, key, value):
    '''
    Функция по внесению изменений в кофигурационный файл.
    :param name_config_file: наименование конфигурационного файла в проекте
    :param key: значение параметра
    :param value: значение параметра
    :return:
    '''
    with open(name_config_file, 'r') as file:
        temp_data = yaml.safe_load(file)
    with open(name_config_file, 'w') as file:
        temp_data[key] = value
        yaml.dump(temp_data, file)

