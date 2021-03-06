#Python Main 
#--Jignesh : Main Function to Execute Mongo Query--

import pymongo
from pymongo import MongoClient
import os
import xlwings as xw
import pdb

#--Import all Mongoquery files--
from DB_QUERY.RegReport.Python.abc_qry import *

@xw.func
def exe_qry(con,ts_folder,qry_nm,db_nm,params):

        try:
                exp = None
                
                #--Connect MongoDB using Connection String--
                if not con:
                        raise Exception ('Connection string is empty')
                client = MongoClient(con)
                client.server_info()

                #--Convert String to Dictionary--
                s = params.replace("{" ,"");
                finalstring = s.replace("}" , "");
                list = finalstring.split("@@")
                dict = {}
                for i in list:
                        if i.find(":") == -1:
                                raise Exception('Invalid Key value pair in string')            
                        else:
                                keyvalue = i.split("::")
                                m= keyvalue[0].strip('\'')
                                m = m.replace("\"", "")
                                dict[m] = keyvalue[1].strip('"\'')


                #--Execute Mongo Query--
                                
                      
                if qry_nm == "get_abc_msg":       
                        cursor,result = get_abc_msg(client,db_nm,dict)
                        #pdb.set_trace()  
                        print (result)
                        

                else:

                        #--Convert cursor to key:value string--
                        col = []
                        val = []
                        data = "{"
                
                        for doc in cursor:
                                for k, v in doc.items():
                                        col.append(k)
                                        val.append(v)
                        
                        j = len(col)
                        print(j)
                
                        i = 0
                        for i in range(j):
                                data = data + str(col[i]) + ":" + str(val[i]) + "##"                        
                                if i > j:
                                        break
                        
                        data = data[:-3]
                        data = data + "}"

                #print(data)
                                                
                #--Save data to .jason--
                file_path = os.path.join(ts_folder,qry_nm + ".jason")
                f=open(file_path, 'w')
                f.write("%s\n" % str(data))
                f.close()

                if (os.stat(file_path).st_size) == 0:
                        raise Exception('Query returned 0 rows')
                else:
                        raise Exception('SUCCESS')

        except pymongo.errors.ServerSelectionTimeoutError as e:
                #--Connection issue--
                exp = e
                
        except Exception as e:
                exp = e
                
        finally:
            #cursor.close()
            if exp:
                #--Write Log-- 
                log_path = os.path.join(ts_folder,qry_nm + "_Log.txt")
                f=open(log_path, 'w')
                f.write("%s\n" % str(exp))
                f.close()
                print(exp)
