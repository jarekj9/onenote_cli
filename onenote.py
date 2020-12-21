#!/usr/bin/env python3

import requests
import json
import pyperclip
import atexit
import os
import shelve
import logging
import time
import datetime
import argparse
from bs4 import BeautifulSoup
from msal import PublicClientApplication, SerializableTokenCache
from dateutil import parser

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
file_handler = logging.FileHandler('logfile.log')
formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(name)s : %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)


class OneNoteDownload:
    '''Can download notebook'''
    def __init__(self, username):
        self.CLIENT_ID = '1f511e95-ec2f-49b9-a52d-0f164d091f05'
        self.AUTHORITY = 'https://login.microsoftonline.com/common'
        self.URL_SECTIONS = f'https://graph.microsoft.com/v1.0/users/{username}/onenote/sections'
        self.URL_PAGES = f'https://graph.microsoft.com/v1.0/users/{username}/onenote/pages'

        access_token = self.get_access_token()
        self.headers = {'Authorization': f'{access_token}'}
        self.section_ids = self._get_sections_ids()

    def get_access_token(self):
        '''Get access token from cache or request new'''
        cache = SerializableTokenCache()
        if os.path.exists('token_cache.bin'):
            cache.deserialize(open('token_cache.bin', 'r').read())
        if cache.has_state_changed:
            atexit.register(lambda: open('token_cache.bin', 'w').write(cache.serialize()))

        app = PublicClientApplication(self.CLIENT_ID, authority=self.AUTHORITY, token_cache=cache)

        token_response = None
        accounts = app.get_accounts()
        if accounts:
            print("Pick the account you want to use to proceed:")
            for index, account in enumerate(accounts):
                print(index, account["username"])
            account_nr = int(input("Type number: "))
            chosen = accounts[account_nr]
            token_response = app.acquire_token_silent(["Notes.Read"], account=chosen)

        if not token_response:
            print('Trying to get token...')
            flow = app.initiate_device_flow(scopes=["Notes.Read"])
            print(flow['message'])
            if 'enter the code ' in flow['message']:
                auth_code = flow['message'].split('enter the code ')[1].split()[0]
                pyperclip.copy(auth_code)
                print(f'Code {auth_code} has been copied to clipboard.')

            token_response = app.acquire_token_by_device_flow(flow)
        if "access_token" in token_response:
            return token_response["access_token"]
        else:
            print(token_response.get("error"))
            print(token_response.get("error_description"))
            print(token_response.get("correlation_id"))

    def _get_sections_ids(self):
        '''Returns dict with mapping: {section name: section id}'''
        section_ids = {}
        next_link = self.URL_SECTIONS

        while next_link:
            for attempt in range(3):
                try:
                    resp = requests.get(next_link, headers=self.headers)
                    sections_response = json.loads(resp.text)
                    for section_data in sections_response['value']:
                        name = section_data["displayName"]
                        if name in section_ids:  # if sections are duplicated
                            existing_date = parser.isoparse(section_ids[name]["lastModifiedDateTime"])
                            new_date = parser.isoparse(section_data["lastModifiedDateTime"])
                            if new_date > existing_date:
                                section_ids[name] = section_data
                        else:
                            section_ids[name] = section_data
                except Exception as e:
                    logger.warning(e)
                    logger.warning(resp.text)
                    logger.warning(f'Retrying {attempt} time.')
                    time.sleep(2)

            next_link = sections_response.get('@odata.nextLink')

        return section_ids

    def get_pages(self, section_name):
        '''Returns dict with mapping: {page title: page id}'''
        section_id = self.section_ids.get(section_name).get('id')
        pages = {}
        next_link = f'{self.URL_SECTIONS}/{section_id}/pages'

        while next_link:
            print(f'Reading next link: {next_link}')

            for attempt in range(3):
                try:
                    resp = requests.get(next_link, headers=self.headers)
                    pages_response = json.loads(resp.text)
                    for pages_data in pages_response.get('value'):
                        pages[pages_data.get("title")] = pages_data.get("id")
                    break
                except Exception as e:
                    logger.warning(e)
                    logger.warning(resp.text)
                    logger.warning(f'Retrying {attempt} time.')
                    time.sleep(2)

            next_link = pages_response.get('@odata.nextLink')

        return pages

    def get_note_text(self, note_id):
        '''Gets full note with html'''
        resp = requests.get(f'{self.URL_PAGES}/{note_id}/content', headers=self.headers)
        soup = BeautifulSoup(resp.text, features='lxml')
        return soup.text

    def get_note(self, page_id):
        '''Gets text without html from a note'''
        resp = requests.get(f'{self.URL_PAGES}/{page_id}/content', headers=self.headers)
        return resp.text


class OneNoteOffline:
    '''For reading offline data'''
    def __init__(self):
        with shelve.open('shelve.lib') as lib:
            self.notes = dict(lib)

    def _find_titles_with_keyword(self, section, keyword):
        return [title for title in self.notes[section].keys() if keyword.casefold() in title.casefold()]

    def _find_titles_with_keyword_in_page(self, section, keyword):
        return [title for title, page in self.notes[section].items() if keyword.casefold() in page.casefold()]

    def display(self):
        '''Display notes according to given parameters'''

        # finds titles by keyword in page content
        if args.find:
            print('Following titles have  been found:')
            for section_name in self.notes:
                titles_with_keyword_in_page = self._find_titles_with_keyword_in_page(section_name, args.find)
                if titles_with_keyword_in_page:
                    print(f'##### SECTION: {section_name} #####')
                    for title in titles_with_keyword_in_page:
                        print(f'          TITLE: {title}')

        # shows all sections
        elif args.allsections:
            print('\n'.join(self.notes.keys()))

        # shows all titles in section
        elif args.alltitles and args.section:
            print('\n'.join(self.notes[args.section].keys()))

        # finds titles with keyword
        elif args.title:
            print('Following titles have been found:')
            for section_name in self.notes:
                titles_with_keyword = self._find_titles_with_keyword(section_name, args.title)
                if titles_with_keyword:
                    print(f'##### SECTION: {section_name} #####')
                    for title in titles_with_keyword:
                        print(f'          TITLE: {title}')

        # shows specific note
        if args.section and args.title:
            titles_with_keyword = self._find_titles_with_keyword(args.section, args.title)
            print(titles_with_keyword)
            for title in titles_with_keyword:
                print(f'          TITLE: {title} ##########')
                soup = BeautifulSoup(self.notes[args.section][title], features='lxml')
            print(soup.text)


if __name__ == '__main__':

    arg_parser = argparse.ArgumentParser(
        description='Can download onenote notebook for specific account and read its contents after.'
    )
    arg_parser.add_argument('--user', '-u', help='Provide login like: xxxx@outlook.com to download notebook.')
    arg_parser.add_argument(
        '--title',
        '-t', default=False,
        help='Use alone to find keyword in titles, or with -s to show note from specific section.'
    )
    arg_parser.add_argument('--section', '-s', default=False, help='Use with -t or with --alltitles')
    arg_parser.add_argument('--find', '-f', default=False, help='Finds titles by keyword in page content')
    arg_parser.add_argument(
        "--alltitles",
        action="store_true",
        help='Show all titles in specific section (use with -s)'
    )
    arg_parser.add_argument("--allsections", action="store_true", help='Show all sections')
    args = arg_parser.parse_args()

    if args.user:
        started_at = time.monotonic()
        onenote = OneNoteDownload(args.user)
        for section_name, section_id in onenote.section_ids.items():
            print(section_name, section_id)
            pages = onenote.get_pages(section_name)
            all_section_notes = {}
            for title, page_id in pages.items():
                print(f'Reading page: {title}')
                all_section_notes[title] = onenote.get_note(page_id)
                with shelve.open('shelve.lib') as lib:
                    lib[section_name] = all_section_notes

        seconds_taken = int(time.monotonic() - started_at)
        time_taken = datetime.timedelta(seconds=seconds_taken)
        print(f'Finished in {time_taken} minutes')

    else:
        onenote_offline = OneNoteOffline()
        onenote_offline.display()
