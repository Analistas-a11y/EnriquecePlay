import unittest

from app import (
    DEFAULT_ENRICHMENT_CONCURRENCY,
    MAX_ENRICHMENT_CONCURRENCY,
    get_enrichment_concurrency,
)


class EnrichmentConcurrencyTests(unittest.TestCase):
    def test_caps_configured_concurrency_at_provider_limit(self):
        self.assertEqual(
            get_enrichment_concurrency(str(MAX_ENRICHMENT_CONCURRENCY + 60)),
            MAX_ENRICHMENT_CONCURRENCY,
        )

    def test_uses_default_for_missing_or_invalid_value(self):
        self.assertEqual(get_enrichment_concurrency(""), DEFAULT_ENRICHMENT_CONCURRENCY)
        self.assertEqual(get_enrichment_concurrency("invalid"), DEFAULT_ENRICHMENT_CONCURRENCY)

    def test_keeps_valid_concurrency_and_rejects_zero(self):
        self.assertEqual(get_enrichment_concurrency("12"), 12)
        self.assertEqual(get_enrichment_concurrency("0"), 1)


if __name__ == "__main__":
    unittest.main()
