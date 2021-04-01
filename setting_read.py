import configparser

def read_config(section):
    # coding: utf-8
    # --------------------------------------------------
    # iniファイルの読み込み
    # --------------------------------------------------
    config_ini = configparser.ConfigParser()
    config_ini.read('config.ini', encoding='utf-8')

    return config_ini[section]
