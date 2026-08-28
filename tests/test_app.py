import unittest

from ahu_utils import extract_ahu_number


class AhuNumberTest(unittest.TestCase):
    def test_normalizes_ocr_value(self):
        self.assertEqual(extract_ahu_number('공 조 기 - 37'), '37')
        self.assertEqual(extract_ahu_number('AHU-42'), '42')

    def test_falls_back_to_filename(self):
        filename = '/tmp/uuid_AHU-37_air_change_rate.pdf'
        self.assertEqual(extract_ahu_number('unknown', filename), '37')

    def test_rejects_zero_and_falls_back_to_filename(self):
        filename = '/tmp/uuid_AHU-33_airborne_particle.pdf'
        self.assertEqual(extract_ahu_number('0', filename), '33')
        self.assertEqual(extract_ahu_number('AHU-0'), 'unknown')

    def test_filename_ahu_overrides_an_incorrect_ocr_number(self):
        filename = '/tmp/uuid_AHU-33_hepa_filter.pdf'
        self.assertEqual(extract_ahu_number('1', filename), '33')


if __name__ == '__main__':
    unittest.main()
