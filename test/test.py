import unittest
from unittest.mock import patch
from unittest import mock
from onenote import OneNoteDownload


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
