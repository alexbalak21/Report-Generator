import unittest
from app.core.processors import apply_operations
import datetime


class TestProcessors(unittest.TestCase):
    def test_apply_operations_numeric_pipeline(self):
        value = 0.738026819923372
        result = apply_operations(value, [
            {"type": "multiply", "value": 100},
            {"type": "round", "decimals": 0},
            {"type": "suffix", "value": "%"}
        ])
        self.assertEqual(result, "74%")

    def test_apply_operations_arithmetic(self):
        result = apply_operations(10, [
            {"type": "add", "value": 5},
            {"type": "multiply", "value": 2},
            {"type": "subtract", "value": 3},
            {"type": "divide", "value": 4},
            {"type": "round", "decimals": 1}
        ])
        self.assertEqual(result, 6.8)

    def test_apply_operations_string_transform(self):
        result = apply_operations(" Hello ", [
            {"type": "strip"},
            {"type": "upper"},
            {"type": "suffix", "value": "!"}
        ])
        self.assertEqual(result, "HELLO!")

    def test_apply_operations_date_format_object(self):
        today = datetime.date(2026, 6, 29)
        result = apply_operations(today, [
            {"type": "date_format", "format": "%d/%m/%Y"}
        ])
        self.assertEqual(result, "29/06/2026")

    def test_apply_operations_date_format_string(self):
        result = apply_operations("2026-06-21", [
            {"type": "date_format", "format": "%d/%m/%Y"}
        ])
        self.assertEqual(result, "21/06/2026")


if __name__ == "__main__":
    unittest.main()
