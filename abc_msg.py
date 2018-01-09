#Sample Qry File
def get_abc_msg(client,db_nm,dict) :

    #pdb.set_trace()
    from pymongo import MongoClient

    #--Get query parameters from  dictionary--
    a = dict.get("a")
    b = dict.get("b")
    c = dict.get("c")
   
      #--Connect DB--
    database = client[db_nm]

    #--Connect Collection--
    collection = database.TransactionData

    #--Any MongoQuery --
    query = {}
    query["aa"] = a
    query["bb"] = b
    query["c.e.x"] = c
    
    projection = {}
    projection["j"] = 1
    
    cursor = collection.find(query, projection = projection)
	
    return cursor
