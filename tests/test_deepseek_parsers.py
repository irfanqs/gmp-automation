import unittest

from deepseek_ocr.parsers import parse_air_change_rate


class AirChangeRateParserTest(unittest.TestCase):
    def test_parses_spaced_metadata_and_groups_measurement_rows(self):
        page = """
        <table>
          <tr><td>측 정 일 자</td><td>2025. 03. 27</td><td>해 당 공 조 기</td><td>공 조 기 − 37</td></tr>
          <tr><td>측 정 결 과</td><td>적 합</td></tr>
        </table>
        <table>
          <tr>
            <th rowspan="2">NO.</th><th rowspan="2">청정 등급</th><th rowspan="2">실번호</th>
            <th rowspan="2">실명</th><th rowspan="2">체적 (m3)</th><th rowspan="2">측정 번호</th>
            <th colspan="2">측정값</th>
          </tr>
          <tr><th>풍량 (m3/hr)</th><th>환기횟수 (회/hr)</th></tr>
          <tr>
            <td rowspan="3">1</td><td rowspan="3">D</td><td rowspan="3">3101</td>
            <td rowspan="3">칭량실</td><td rowspan="3">50.0</td><td>1</td><td>480.2</td><td rowspan="3">20</td>
          </tr>
          <tr><td>2</td><td>500.3</td></tr>
          <tr><td>합계</td><td>980.5</td></tr>
        </table>
        """

        result = parse_air_change_rate([page])

        self.assertEqual(result['ahu'], '37')
        self.assertEqual(result['date'], '2025.03.27')
        self.assertEqual(result['result'], '적합')
        self.assertEqual(len(result['rooms']), 1)
        room = result['rooms'][0]
        self.assertEqual(room['room_number'], '3101')
        self.assertEqual(room['total_air_flow'], 980.5)
        self.assertEqual(room['ach'], 20)
        self.assertEqual(
            room['air_flow_measurements'],
            [{'point': 1, 'air_flow': 480.2}, {'point': 2, 'air_flow': 500.3}],
        )


if __name__ == '__main__':
    unittest.main()
