import unittest

from plant_parser_production import (
    RULE_BY_KEY,
    Observation,
    display_observation,
    parse_container_size,
    parse_explicit_tolerance,
    parse_growth_rate,
    parse_hardiness,
    parse_length_values,
    parse_lifespan,
    parse_ph,
    parse_rule_value,
)


class ParserCoreTests(unittest.TestCase):
    def test_length_range_cm_to_m(self):
        values = parse_length_values("150–200 cm")
        self.assertTrue(values)
        self.assertAlmostEqual(values[0].min_value, 1.5)
        self.assertAlmostEqual(values[0].max_value, 2.0)
        self.assertEqual(values[0].canonical_unit, "m")

    def test_length_upper_bound(self):
        values = parse_length_values("до 200 cm")
        self.assertEqual(values[0].normalized_value, "≤2 m")
        self.assertIsNone(values[0].min_value)
        self.assertAlmostEqual(values[0].max_value, 2.0)

    def test_multiple_lengths_preserved(self):
        values = parse_length_values("высота 1.5 m, иногда 2 m")
        self.assertGreaterEqual(len(values), 2)

    def test_display_cm_derived_from_canonical_m(self):
        o = Observation("x", "mature_height", "", "", "150-200 cm", "1.5–2 m", 1.5, 2.0, "cm", "m", 1.0, "test", "", "", "")
        self.assertEqual(display_observation(o, "cm"), "150–200 cm")

    def test_container_dimensions_shared_unit(self):
        values = parse_container_size("60×60 cm")
        self.assertTrue(any(v.normalized_value == "0.6×0.6 m" for v in values))

    def test_container_code_preserved(self):
        values = parse_container_size("C20")
        self.assertTrue(any(v.normalized_value == "C20" for v in values))

    def test_growth_rate_numeric(self):
        values = parse_growth_rate("30 cm per year")
        self.assertTrue(any(v.canonical_unit == "m/year" and abs(v.min_value - 0.3) < 1e-9 for v in values))

    def test_lifespan_range(self):
        values = parse_lifespan("50-100 years")
        self.assertEqual(values[0].min_value, 50)
        self.assertEqual(values[0].max_value, 100)

    def test_wet_soil_not_waterlogging_tolerance(self):
        self.assertEqual(parse_explicit_tolerance("prefers wet soil", "waterlogging"), [])

    def test_explicit_waterlogging(self):
        self.assertTrue(parse_explicit_tolerance("tolerates waterlogging", "waterlogging"))

    def test_waterlogging_label_yes(self):
        rule = RULE_BY_KEY["waterlogging_tolerance"]
        vals = parse_rule_value(rule, "да", "Устойчивость к переувлажнению")
        self.assertEqual(vals[0].normalized_value, "устойчиво / переносит")

    def test_drought_label_high(self):
        rule = RULE_BY_KEY["drought_tolerance"]
        vals = parse_rule_value(rule, "High", "Drought tolerance")
        self.assertEqual(vals[0].normalized_value, "высокая")

    def test_hardy_does_not_invent_usda(self):
        values = parse_hardiness("very hardy")
        self.assertTrue(values)
        self.assertFalse(any("USDA" in v.normalized_value for v in values))

    def test_explicit_usda(self):
        values = parse_hardiness("USDA zone 4-6")
        self.assertTrue(any("USDA 4–6" == v.normalized_value for v in values))

    def test_ph_range(self):
        values = parse_ph("pH 5.5-7.0")
        self.assertTrue(values)
        numeric = [v for v in values if v.min_value is not None]
        self.assertAlmostEqual(numeric[0].min_value, 5.5)
        self.assertAlmostEqual(numeric[0].max_value, 7.0)

    def test_container_suitability_label_yes(self):
        rule = RULE_BY_KEY["in_container"]
        vals = parse_rule_value(rule, "да", "Выращивание в контейнере")
        self.assertEqual(vals[0].normalized_value, "да")

    def test_container_size_is_not_suitability(self):
        rule = RULE_BY_KEY["in_container"]
        vals = parse_rule_value(rule, "Container size C5", "Container size")
        self.assertEqual(vals, [])


if __name__ == "__main__":
    unittest.main()
