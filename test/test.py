#!/usr/bin/env python3

import sys
import io
import shelve
import logging
import unittest
from unittest.mock import patch
from unittest import mock
import os.path
sys.path.append(
    os.path.abspath(os.path.join(os.path.dirname(__file__), os.path.pardir))
)
from onenote import OneNoteDownload, OneNoteOffline

logging.basicConfig(stream=sys.stderr, level=logging.DEBUG)
log = logging.getLogger("TestLog")

class TestOneNoteDownload(unittest.TestCase):

    @patch('onenote.requests.get')
    @patch('onenote.OneNoteDownload.get_access_token')
    def test_get_sections_data(self,
                               get_access_token_mock,
                               requests_get_mock):
        with open('sections_list_response_fixture1.txt') as f:
            sections_fixture_part1 = f.read()
        with open('sections_list_response_fixture2.txt') as f:
            sections_fixture_part2 = f.read()

        mockresponse1 = mock.Mock()
        mockresponse1.text = sections_fixture_part1
        mockresponse2 = mock.Mock()
        mockresponse2.text = sections_fixture_part2
        requests_get_mock.side_effect = [
            mockresponse1,
            mockresponse1,
            mockresponse2,
            mockresponse2
        ]

        onenote = OneNoteDownload('test@outlook.com')
        for section_number in range(1,7):
            assert(onenote.section_data.get(f'SECTION NAME {section_number}'))

    @patch('onenote.OneNoteDownload._get_sections_data')
    @patch('onenote.requests.get')
    @patch('onenote.OneNoteDownload.get_access_token')
    def test_get_pages(self,
                       get_access_token_mock,
                       requests_get_mock,
                       get_sections_data_mock):
        onenote = OneNoteDownload('test@outlook.com')
        onenote.section_data = {
            'SECTION NAME 1':
                             {'id': 'section 1 id'}
        }

        with open('pages_list_response_fixture.txt') as f:
            pages_fixture = f.read()
        mockresponse = mock.Mock()
        mockresponse.text = pages_fixture
        requests_get_mock.return_value = mockresponse

        pages_expected_value = {
            'title1': 'page1_id',
            'title2': 'page2_id',
            'title3': 'page3_id'
        }
        pages = onenote.get_pages('SECTION NAME 1')
        self.assertEqual(pages, pages_expected_value)

    @patch('onenote.OneNoteDownload._get_sections_data')
    @patch('onenote.requests.get')
    @patch('onenote.OneNoteDownload.get_access_token')
    def test_get_note_text(self,
                           get_access_token_mock,
                           requests_get_mock,
                           get_sections_data_mock):
        with open('page_response_fixture.txt') as f:
            page_fixture = f.read()
        mockresponse = mock.Mock()
        mockresponse.text = page_fixture
        requests_get_mock.return_value = mockresponse

        onenote = OneNoteDownload('test@outlook.com')

        expected_note_text = \
        '\n\nPage title\n\n\n\n\n\nHeader\n\nNote text\n\n\n\n\n'

        note_text = onenote.get_note_text('test_id')
        self.assertEqual(note_text, expected_note_text)


class TestOneNoteOffline(unittest.TestCase):
    
    @classmethod
    def setUpClass(cls):
        with shelve.open('shelve_fixture.lib') as lib:
            notes_fixture = dict(lib)
        cls.onenote_offline = OneNoteOffline()
        cls.onenote_offline.notes = notes_fixture

    @patch('sys.stdout', new_callable=io.StringIO)
    def test_display_titles_with_keyword_in_page(self, mock_stdout):
        self.onenote_offline._display_titles_with_keyword_in_page('Page Text11')
        self.assertIn('##### SECTION: Section name1 #####', mock_stdout.getvalue())
        self.assertIn('TITLE: Title11', mock_stdout.getvalue())
        self.assertNotIn('##### SECTION: Section name2 #####', mock_stdout.getvalue())
        self.assertNotIn('TITLE: Title12', mock_stdout.getvalue())

    @patch('sys.stdout', new_callable=io.StringIO)
    def test_print_all_sections(self, mock_stdout):
        self.onenote_offline._print_all_sections()
        self.assertIn('Section name1', mock_stdout.getvalue())
        self.assertIn('Section name2', mock_stdout.getvalue())

    @patch('sys.stdout', new_callable=io.StringIO)
    def test_print_all_titles_in_section(self, mock_stdout):
        self.onenote_offline._print_all_titles_in_section('Section name1')
        self.assertIn('Title11', mock_stdout.getvalue())
        self.assertIn('Title12', mock_stdout.getvalue())
        self.assertNotIn('Title21', mock_stdout.getvalue())
        self.assertNotIn('Title22', mock_stdout.getvalue())

    @patch('sys.stdout', new_callable=io.StringIO)
    def test_print_note(self, mock_stdout):
        self.onenote_offline._print_note('Section name1', 'Title11')
        self.assertIn('Page Text11', mock_stdout.getvalue())
        self.assertNotIn('Page Text12', mock_stdout.getvalue())
        self.assertNotIn('Page Text21', mock_stdout.getvalue())
        self.assertNotIn('Page Text22', mock_stdout.getvalue())

    @patch('sys.stdout', new_callable=io.StringIO)
    def test_print_titles_with_keyword(self, mock_stdout):
        self.onenote_offline._print_titles_with_keyword('Title1')
        self.assertIn('TITLE: Title11', mock_stdout.getvalue())
        self.assertIn('TITLE: Title12', mock_stdout.getvalue())
        self.assertNotIn('Title2', mock_stdout.getvalue())
        self.assertNotIn('Section name2', mock_stdout.getvalue())

    @patch('onenote.OneNoteOffline._print_titles_with_keyword')
    @patch('onenote.OneNoteOffline._print_note')
    @patch('onenote.OneNoteOffline._print_all_titles_in_section')
    @patch('onenote.OneNoteOffline._print_all_sections')
    @patch('onenote.OneNoteOffline._display_titles_with_keyword_in_page')
    def test_display_notes(self,
                           mock_display_titles_with_keyword_in_page,
                           mock_print_all_sections,
                           mock_print_all_titles_in_section,
                           mock_print_note,
                           mock_print_titles_with_keyword):
        args_mock = mock.Mock()
        args_mock.section = False
        args_mock.allsections = False
        args_mock.title = False
        args_mock.alltitles = False

        args_mock.find = 'Text11'
        self.onenote_offline.display_notes(args_mock)
        mock_display_titles_with_keyword_in_page.assert_called_with('Text11')
        args_mock.find = False

        args_mock.allsections = True
        self.onenote_offline.display_notes(args_mock)
        self.assertTrue(mock_print_all_sections.called)
        args_mock.allsections = False

        args_mock.section = 'Section name1'
        args_mock.alltitles = True
        self.onenote_offline.display_notes(args_mock)
        mock_print_all_titles_in_section.assert_called_with('Section name1')
        args_mock.section = False
        args_mock.alltitles = False

        args_mock.section = 'Section name1'
        args_mock.title = 'Title11'
        self.onenote_offline.display_notes(args_mock)
        mock_print_note.assert_called_with('Section name1', 'Title11')
        args_mock.section = False
        args_mock.title = False

        args_mock.title = 'Title11'
        self.onenote_offline.display_notes(args_mock)
        mock_print_titles_with_keyword.assert_called_with('Title11')
        args_mock.title = False


if __name__ == '__main__':
    unittest.main()
