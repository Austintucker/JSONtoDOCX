import pymongo, ssl

uri = "mongodb://localhost:C2y6yDjf5%2FR%2Bob0N8A7Cgv30VRDJIWEHLM%2B4QDU5DE2nQ9nDuVTqobD4b8mGGyPMbIZnqyMsEcaGQy67XIw%2FJw%3D%3D@localhost:10255/?ssl=true"
client = pymongo.MongoClient(uri, ssl_cert_reqs=ssl.CERT_NONE)

db = client['ResilientResumes']

try: 
    db.command("serverStatus")
except Exception as e: 
    print(e)
else: 
    print("You are connected!")

