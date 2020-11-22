import logging
from typing import Dict

logging.basicConfig(level=logging.DEBUG)

import pickle
import os.path
from googleapiclient.discovery import build
import google
import google_auth_oauthlib
import pandas
import calendar

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/contacts']
ZODIAC_FILE_LOC = None
with open("zodiac_file_location.txt") as f:
    ZODIAC_FILE_LOC = f.readline().strip()
zodiac_names_and_birthdays: Dict[str, Dict[str, int]] = dict()
"""{ Person : { day/month/year : int} }"""

contacts_analyzed = 0
contacts_updated = 0
known_birthdays_count = 0


def main():
    pd = pandas.read_excel(ZODIAC_FILE_LOC)
    for idx, row in pd.iterrows():
        zodiac_names_and_birthdays[row["Name"]] = dict()
        zodiac_names_and_birthdays[row["Name"]]["day"] = row["Day"]
        zodiac_names_and_birthdays[row["Name"]]["month"] = list(calendar.month_name).index(
            row["Month"])  # Convert month name to int
        zodiac_names_and_birthdays[row["Name"]]["year"] = row["Year"]

    creds = init_credentials()
    service = build('people', 'v1', credentials=creds, cache_discovery=False)

    page_token = None
    while True:
        # people.connections.list provides a List of authenticated user's contacts
        results = service.people().connections().list(
            resourceName='people/me',
            pageSize=1000,
            pageToken=page_token,
            personFields='names,emailAddresses,nicknames,birthdays').execute()

        connections = results.get('connections', [])
        do_stuff_to_connections_list(service, connections)

        # See if there are more pages to process
        if "nextPageToken" in results:
            page_token = results["nextPageToken"]
        else:
            # Done analyzing all contacts, kill the loop
            break

    logging.info("Total contacts analyzed: " + str(contacts_analyzed))
    logging.info("Total contacts updated: " + str(contacts_updated))
    logging.info("Total known birthdays: " + str(known_birthdays_count))


def do_stuff_to_connections_list(service, connectionsList: list):
    global contacts_updated
    global contacts_analyzed
    global known_birthdays_count

    for person in connectionsList:
        # Try to get a Person's name.
        # No name usually means I just have an email for this contact. Skip them.
        names = person.get('names', [])
        if not names:
            continue
        else:
            name = names[0].get('displayName').strip()

        # Attempt to get the person's birthday
        birthdays = person.get('birthdays', [])
        if birthdays:
            birthday = birthdays[0].get('date')

        if birthdays and names:
            try:
                logging.info("NAME: " + name + "\t\tBirthday: " + str(birthday["month"]) + "/" + str(
                    birthday["day"]) + "/" + str(birthday["year"]))
            except Exception as e:
                logging.debug("NAME:" + name + " caused Exception: " + str(e))

            known_birthdays_count = known_birthdays_count + 1

        elif names:
            # I have a name but no birthday for this person. These are the ones we want to fix.
            if name in zodiac_names_and_birthdays.keys():
                logging.debug("Adding birthday (" + str(zodiac_names_and_birthdays[name]["month"]) + "/" + str(
                    zodiac_names_and_birthdays[name]["day"]) + "/" + str(
                    zodiac_names_and_birthdays[name]["year"]) + ") for " + name)

                # logging.debug("BEFORE:")
                # logging.debug(str(person))

                person["birthdays"] = list()
                birthday_obj = dict()
                birthday_obj["metadata"] = dict()
                birthday_obj["metadata"]["primary"] = True
                birthday_obj["metadata"]["verified"] = True
                birthday_obj["date"] = dict()
                birthday_obj["text"] = str(zodiac_names_and_birthdays[name]["month"]) + "/" + str(
                    zodiac_names_and_birthdays[name]["day"]) + "/" + str(zodiac_names_and_birthdays[name]["year"])
                birthday_obj["date"]["year"] = zodiac_names_and_birthdays[name]["year"]
                birthday_obj["date"]["month"] = zodiac_names_and_birthdays[name]["month"]
                birthday_obj["date"]["day"] = zodiac_names_and_birthdays[name]["day"]
                person["birthdays"] = birthday_obj

                # logging.debug("AFTER:")
                # logging.debug(str(person))

                result = service.people().updateContact(
                    resourceName=person["resourceName"],
                    body=person,
                    updatePersonFields='birthdays'
                ).execute()

                contacts_updated = contacts_updated + 1
                known_birthdays_count = known_birthdays_count + 1
            else:
                pass
                # logging.debug(name)
        else:
            logging.debug("names came up empty for some reason.")
            logging.debug(str(person))
        contacts_analyzed = contacts_analyzed + 1


def init_credentials():
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = google_auth_oauthlib.flow.InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    return creds


if __name__ == '__main__':
    main()
