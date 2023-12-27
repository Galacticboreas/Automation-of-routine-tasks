import configparser

# file example.ini
# [DEFAULT]
# ServerAliveInterval = 45
# Compression = yes
# CompressionLevel = 9
# ForwardX11 = yes

# [forge.example]
# User = hg

# [topsecret.server.example]
# Port = 50022
# ForwardX11 = no

config = configparser.ConfigParser()
config.sections()

print(config.read('example.ini'))
# ['example.ini']

print(config.sections())
# ['forge.example', 'topsecret.server.example']

print('forge.example' in config)
# True

print('python.org' in config)
False

print(config['forge.example']['User'])
# hg

print(config['DEFAULT']['Compression'])
# yes

topsecret = config['topsecret.server.example']
print(topsecret['ForwardX11'])
# no

print(topsecret['Port'])
# 50022

for key in config['forge.example']:  
    print(key)
# user
# serveraliveinterval
# compression
# compressionlevel
# forwardx11
    
print(config['forge.example']['ForwardX11'])
# yes
