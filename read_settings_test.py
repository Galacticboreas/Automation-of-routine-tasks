import configparser

config = configparser.ConfigParser()
config.sections()

print(config.read('example.ini'))

print(config.sections())
