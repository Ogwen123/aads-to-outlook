import requests
import json
import os
import uuid
from datetime import datetime

done = []

def get_config_option(option):
    with open("config.json", "r") as config_file:
        config = json.load(config_file)
        return config[option]

URL = "https://ds-api.theaa.digital/graphql"
HEADERS = {
    "authorization": get_config_option("aads.auth")
}
JSON = {
    "operationName":"GetBookedLessons",
    "variables":{"learnerId":get_config_option("aads.id")},
    "query":"query GetBookedLessons($learnerId: ID!) {\n  learner(id: $learnerId) {\n    lessons {\n      id\n      status\n      startDateTime\n      endDateTime\n      __typename\n    }\n    __typename\n  }\n}\n"
    }

def get_data():
    response = requests.post(url=URL, json=JSON, headers=HEADERS)
    if not response.status_code == 200:
        print("Error: {}".format(response.status_code))
        return
    else:
        return json.loads(response.content)["data"]

def done_lessons(method, id=None):
    global done
    try:# make sure the done.json file exists
        temp = open("done.json", "x")
        temp.write("[]")
        temp.close()
    except:
        pass
    if method == "get":
        with open("done.json", "r") as done_file:
            done_temp = json.load(done_file)
            done = done_temp
            return
    elif method == "add":
        with open("done.json", "w") as done_file:
            done_file.write(json.dumps(done))
            return
    else:
        return

def check_for_past_date(start_date):# returns true if in the past
    print(datetime.now().month)
    now = datetime.now()
    date = start_date.split("T")[0]
    year = int(date.split("-")[0])
    month = int(date.split("-")[1])
    day = int(date.split("-")[2])
    if year < now.year:
        return True
    elif year == now.year:
        if month < now.month:
            return True
        elif month == now.month:
            if day < now.day:
                return True
            else:
                return False
        else:
            return False
    else: 
        return False


def add_to_outlook(lesson):
    res = requests.post("https://outlook.office365.com/owa/service.svc?action=CreateCalendarEvent&app=Calendar", headers={
    "Accept": '*/*',
    'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
    "Connection": 'keep-alive',
    "Cookie": get_config_option("outlook.cookies").strip(),
    "DNT": '1',
    "Origin": 'https://outlook.office365.com',
    'Sec-Fetch-Dest': 'empty',
    'Sec-Fetch-Mode': 'cors',
    'Sec-Fetch-Site': 'same-origin',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:109.0) Gecko/20100101 Firefox/119.0',
    "action": 'CreateCalendarEvent',
    'content-type': 'application/json; charset=utf-8',
    "prefer": 'exchange.behavior="IncludeThirdPartyOnlineMeetingProviders"',
    'sec-ch-ua': '"Not/A)Brand";v="99", "Google Chrome";v="115", "Chromium";v="115"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': '"Linux"',
    'x-owa-canary': get_config_option("outlook.cookies").split('X-OWA-CANARY=')[1].split(';')[0],
    'x-req-source': 'Calendar'
    }, json={
    "__type": "CreateItemJsonRequest:#Exchange",
    "Header": {
        "__type": "JsonRequestHeaders:#Exchange",
        "RequestServerVersion": "V2018_01_08",
        "TimeZoneContext": {
            "__type": "TimeZoneContext:#Exchange",
            "TimeZoneDefinition": {
                "__type": "TimeZoneDefinitionType:#Exchange",
                "Id": "GMT Standard Time"
            }
        }
    },
    "Body": {
        "__type": "CreateItemRequest:#Exchange",
        "Items": [
            {
                "__type": "CalendarItem:#Exchange",
                "FreeBusyType": "Busy",
                "ParentFolderId": {
                    "Id": get_config_option("outlook.calendar_id"),
                    "mailboxInfo": {
                        "mailboxRank": "Coprincipal",
                        "mailboxSmtpAddress": get_config_option("outlook.email"),
                        "sourceId": f'main:m365:{get_config_option("outlook.email")}',
                        "type": "UserMailbox",
                        "userIdentity": get_config_option("outlook.email")
                    }
                },
                "Sensitivity": "Normal",
                "Subject": "Driving Lesson",
                "Body": {
                    "BodyType": "HTML",
                    "Value": f"<div style=\"font-family: Calibri, Arial, Helvetica, sans-serif; font-size: 12pt; color: rgb(0, 0, 0);\" class=\"elementToProof\">Driving Lesson</div>"
                },
                "Start": lesson["startDateTime"],
                "End": lesson["endDateTime"],
                "IsAllDayEvent": False,
                "ReminderMinutesBeforeStart": 5, # change this if u want ig
                "ReminderIsSet": True,
                "CharmId": 57,
                "MuteNotifications": False,
                "CalendarEventClassifications": [],
                "Resources": [],
                "Locations": [
                    {
                        "Id": "Home",
                        "DisplayName": "Home"
                    }
                ],
                "IsDraft": False,
                "DoNotForwardMeeting": False,
                "IsResponseRequested": True,
                "StartTimeZoneId": "GMT Standard Time",
                "EndTimeZoneId": "GMT Standard Time",
                "HideAttendees": False,
                "AppendOnSend": [],
                "PrependOnSend": [],
                "IsBookedFreeBlocks": False,
                "AssociatedTasks": [],
                "CollabSpace": None,
                "DocLinks": [],
                "ItemId": {
                    "__type": "ItemId:#Exchange",
                    "Id": str(uuid.uuid4()),
                    "mailboxInfo": {
                        "mailboxRank": "Coprincipal",
                        "mailboxSmtpAddress": get_config_option("outlook.email"),
                        "sourceId": f'main:m365:{get_config_option("outlook.email")}',
                        "type": "UserMailbox",
                        "userIdentity": get_config_option("outlook.email")
                    }
                },
                "EffectiveRights": {
                    "Read": True,
                    "Modify": True,
                    "Delete": True,
                    "ViewPrivateItems": True
                },
                "IsOrganizer": True,
                "ExtendedProperty": []
            }
        ],
        "SavedItemFolderId": {
            "__type": "TargetFolderId:#Exchange",
            "BaseFolderId": {
                "__type": "FolderId:#Exchange",
                "Id": get_config_option("outlook.calendar_id"),
            }
        },
        "ClientSupportsIrm": True,
        "UnpromotedInlineImageCount": 0,
        "ItemShape": {
            "__type": "ItemResponseShape:#Exchange",
            "BaseShape": "IdOnly"
        },
        "SendMeetingInvitations": "SendToNone",
        "OutboundCharset": "AutoDetect",
        "UseGB18030": False,
        "UseISO885915": False
    }
    })
    print(res.status_code)

def main():
    done_lessons("get")
    lessons = get_data()["learner"]["lessons"]

    for lesson in lessons:
        if not lesson["status"] == "Cancelled" and not lesson["id"] in done and not check_for_past_date(lesson["startDateTime"]):
            done.append(lesson["id"])
            print(lesson)
            add_to_outlook(lesson)
    
    done_lessons("add")

if __name__ == "__main__":
    main()