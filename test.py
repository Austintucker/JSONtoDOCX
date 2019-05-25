import json, os

resumeDirectory = f"{os.getcwd()}/JSON_Clean/"

addressDoc  = f"{resumeDirectory}/address.json"
resumeDoc   = f"{resumeDirectory}/resume.json"
userDoc     = f"{resumeDirectory}/user.json"

with open(addressDoc, "r") as read_file:
    address = json.load(read_file)
with open(resumeDoc, "r") as read_file:
    resume = json.load(read_file)
with open(userDoc, "r") as read_file:
    user = json.load(read_file)


