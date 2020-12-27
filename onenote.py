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


class OneNoteDownload:
    '''Can download notebook'''
    def __init__(self, username):
        self.CLIENT_ID = '1f511e95-ec2f-49b9-a52d-0f164d091f05'
        self.AUTHORITY = 'https://login.microsoftonline.com/common'
        self.URL_SECTIONS = f'https://graph.microsoft.com/v1.0/users/{username}/onenote/sections'
        self.URL_PAGES = f'https://graph.microsoft.com/v1.0/users/{username}/onenote/pages'
        self.logger = setup_logger()
        access_token = self.get_access_token()
        self.headers = {'Authorization': f'{access_token}'}
        self.section_data = self._get_sections_data()

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

    def _get_sections_data(self):
        '''Returns dict with mapping: {section name: section data}'''
        sections_data = {}
        next_link = self.URL_SECTIONS

        while next_link:
            for attempt in range(3):
                try:
                    sections_data.update(self._get_sections_data_from_link(next_link))
                except Exception as e:
                    self.logger.warning(e)
                    self.logger.warning(f'Retrying {attempt} time.')
                    time.sleep(2)

            sections_dict = self._load_response_json(next_link)
            next_link = sections_dict.get('@odata.nextLink')

        return sections_data
    
    def _get_sections_data_from_link(self, link):
        sections_data = {}
        sections = self._load_response_json(link)
        for section in sections['value']:
            section_name = section["displayName"]
            if section_name in sections_data:  # if sections are duplicated
                existing_date = parser.isoparse(sections_data[section_name]["lastModifiedDateTime"])
                new_date = parser.isoparse(section["lastModifiedDateTime"])
                if new_date > existing_date:
                    sections_data[section_name] = section
            else:
                sections_data[section_name] = section
        return sections_data

    def _load_response_json(self, link):
        resp = requests.get(link, headers=self.headers)
        return json.loads(resp.text)

    def get_pages(self, section_name):
        '''Returns dict with mapping: {page title: page id}'''
        section_id = self.section_data.get(section_name).get('id')
        all_pages = {}
        next_link = f'{self.URL_SECTIONS}/{section_id}/pages'

        while next_link:
            print(f'Reading next link: {next_link}')

            for attempt in range(3):
                try:
                    pages_response = self._get_pages_from_link(next_link)
                    all_pages.update(pages_response)
                    break
                except Exception as e:
                    self.logger.warning(e)
                    self.logger.warning(resp.text)
                    self.logger.warning(f'Retrying {attempt} time.')
                    time.sleep(2)
            next_link = pages_response.get('@odata.nextLink')

        return all_pages

    def _get_pages_from_link(self, link):
        pages = {}
        resp = requests.get(link, headers=self.headers)
        pages_response = json.loads(resp.text)
        for pages_data in pages_response.get('value'):
            pages[pages_data.get("title")] = pages_data.get("id")
        return pages

    def get_note_text(self, note_id):
        '''Not used yet'''
        html = self.get_note_html(note_id)
        soup = BeautifulSoup(html, features='lxml')
        return soup.text

    def get_note_html(self, note_id):
        '''Not used yet'''
        resp = requests.get(f'{self.URL_PAGES}/{note_id}/content', headers=self.headers)
        with open('page.txt', 'w') as f:
            f.write(resp.text)
        return resp.text


class OneNoteOffline:
    '''For reading offline data'''
    def __init__(self):
        with shelve.open('shelve.lib') as lib:
            self.notes = dict(lib)

    def _find_titles_with_keyword(self, section, keyword):
        return [
            title for title in self.notes[section].keys()
            if keyword.casefold() in title.casefold()
        ]

    def _find_titles_with_keyword_in_page(self, section, keyword):
        return [
            title for title, page in self.notes[section].items()
            if keyword.casefold() in page.casefold()
        ]

    def _find_sections_with_keyword(self, keyword):
        return [
            section_name for section_name in self.notes.keys()
            if keyword.casefold() in section_name.casefold()
        ]

    def _display_titles_with_keyword_in_page(self, keyword):
        print('Following titles have  been found:')
        for section_name in self.notes:
            titles_with_keyword_in_page = self._find_titles_with_keyword_in_page(section_name, args.find)
            if titles_with_keyword_in_page:
                print(f'##### SECTION: {section_name} #####')
                for title in titles_with_keyword_in_page:
                    print(f'          TITLE: {title}')

    def _print_all_sections(self):
        print('\n'.join(self.notes.keys()))

    def _print_all_titles_in_section(self, section_name):
        print('\n'.join(self.notes[args.section].keys()))

    def _print_note(self, section_name, title):
        '''Shows specific note. If more sections match it only shows
        sections. If one sections, but more titles match, it shows all
        notes with matching titles
        '''
        sections = self._find_sections_with_keyword(section_name)
        if len(sections) == 0:
            print('Provided section name not found.')
        elif len(sections) == 1:
            section_name = sections[0]
            titles_with_keyword = self._find_titles_with_keyword(section_name, title)
            print(f'##### SECTION: {section_name} #####')
            for title in titles_with_keyword:
                print(f'##### TITLE: {title}')
                soup = BeautifulSoup(self.notes[section_name][title], features='lxml')
                print(soup.text)
        elif len(sections) > 1:
            print(f'Section name : {section_name}, matches more than one section: {sections}.')

    def _print_titles_with_keyword(self, keyword):
        '''Finds titles with keyword, but if only one is found, it
        shows note
        '''
        print('Following titles have been found in all sections:')
        found = []
        for section_name in self.notes:
            titles_with_keyword = self._find_titles_with_keyword(section_name, keyword)
            if titles_with_keyword:
                print(f'##### SECTION: {section_name} #####')
                for title in titles_with_keyword:
                    found.append({'section_name': section_name, 'title': title})
                    print(f'          TITLE: {title}')

        if len(found) == 1:
            title = found[0]['title']
            section = found[0]['section_name']
            soup = BeautifulSoup(self.notes[section][title], features='lxml')
            print(soup.text)

    def display_notes(self):
        if args.find:
            self._display_titles_with_keyword_in_page(args.find)
        elif args.allsections:
            self._print_all_sections()
        elif args.section and args.alltitles:
            self._print_all_titles_in_section(args.section)
        if args.section and args.title:
            self._print_note(args.section, args.title)
        elif args.title:
            self._print_titles_with_keyword(args.title)

def setup_logger():
    logger = logging.getLogger(__name__)
    logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler('logfile.log')
    formatter = logging.Formatter('%(asctime)s : %(levelname)s : %(name)s : %(message)s')
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    return logger

def parse_arguments():
    arg_parser = argparse.ArgumentParser(
        description='Can download onenote notebook for specific account and read its contents after.'
    )
    arg_parser.add_argument('--user', '-u', help='Provide login like: xxxx@outlook.com to download notebook.')
    arg_parser.add_argument(
        '--title',
        '-t',
        default=False,
        help='Find keyword in titles and display the page (if there is only one matching). '
            +'Can use with together with -s to show page from specific section.'
    )
    arg_parser.add_argument(
        '--section',
        '-s',
        default=False,
        help='Section can be provided together with with -t to print specific page '
            +'or with --alltitles to print all titles in specific section'
    )
    arg_parser.add_argument(
        '--find',
        '-f',
        default=False,
        help='Finds all titles from all sections by keyword in page content'
    )
    arg_parser.add_argument(
        "--alltitles",
        action="store_true",
        help='Show all titles in specific section (use with -s)'
    )
    arg_parser.add_argument("--allsections", action="store_true", help='Show all sections')

    return arg_parser.parse_args()


if __name__ == '__main__':

    args = parse_arguments()

    if args.user:
        started_at = time.monotonic()
        onenote = OneNoteDownload(args.user)
        for section_name, section_dict in onenote.section_data.items():
            print(f"Reading section: {section_name}, {section_dict['id']}")
            pages = onenote.get_pages(section_name)
            all_section_notes = {}
            for title, page_id in pages.items():
                print(f'Reading page: {title}')
                all_section_notes[title] = onenote.get_note_html(page_id)
                with shelve.open('shelve.lib') as lib:
                    lib[section_name] = all_section_notes

        seconds_taken = int(time.monotonic() - started_at)
        time_taken = datetime.timedelta(seconds=seconds_taken)
        print(f'Finished in {time_taken} minutes')

    else:
        OneNoteOffline().display_notes()
