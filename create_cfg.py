import os
import configparser

clientId = os.getenv('azure_clientId')
clientSecret = os.getenv('azure_clientSecret')
tenantId = os.getenv('azure_tenantId')


config = configparser.ConfigParser()
# Add the structure to the file we will create
config.add_section('azure')
config.set('azure', 'clientId', clientId)
config.set('azure', 'clientSecret', clientSecret)
config.set('azure', 'tenantId', tenantId)
config.set('azure', 'authTenant', 'common')
config.set('azure', 'graphUserScopes', 'User.Read Mail.Read Mail.Send')

# Write the new structure to the new file
with open("./config.dev.cfg", 'w') as configfile:
    config.write(configfile)