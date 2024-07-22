from configparser import ConfigParser
def config2(filename='myconn.ini', section='postgresql'):
    # create a parser
    parser = ConfigParser()
    # read config file
    parser.read(filename)
 
    # get section, default to postgresql
    db = {}
    
     
    
    if parser.has_section(section):
        params = parser.items(section)
        for param in params:            
            db[param[0]] = param[1]
            db["client_encoding"] = "utf-8"
    else:
        raise Exception('Section {0} not found in the {1} file'.format(section, filename))
      
    db["database"] = "reporting"
    db["host"] = "rsusdb47.yardispectrum.com"
    db["user"] = "xiaobinz"
    db["port"] = 5435
    db["password"] = "hKB$ud#^05^F3g3"
     
    return db
 
 
