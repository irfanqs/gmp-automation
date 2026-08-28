import unittest

from ahu_utils import extract_ahu_number


class AhuNumberTest(unittest.TestCase):
    def test_normalizes_ocr_value(self):
        self.assertEqual(extract_ahu_number('공 조 기 - 37'), '37')
        self.assertEqual(extract_ahu_number('AHU-42'), '42')

    def test_falls_back_to_filename(self):
        filename = '/tmp/uuid_AHU-37_air_change_rate.pdf'
        self.assertEqual(extract_ahu_number('unknown', filename), '37')


if __name__ == '__main__':
    unittest.main()
