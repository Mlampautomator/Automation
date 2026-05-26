#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Unit tests for porownaj_ceny.py — pure functions only (no network)."""

import pytest
from bs4 import BeautifulSoup

from porownaj_ceny import (
    parse_price,
    _extract_catalog_num,
    _name_similarity,
    _match_confidence,
    _model_anchors,
    _build_search_queries,
    select_offers,
    _extract_product_id,
    _extract_shop_name,
)


# ── parse_price ───────────────────────────────────────────────────────────────

class TestParsePrice:
    def test_simple_integer(self):
        assert parse_price("299 zl") == 299.0

    def test_decimal_comma(self):
        assert parse_price("1 299,99 zl") == 1299.99

    def test_decimal_dot(self):
        assert parse_price("1299.99") == 1299.99

    def test_nbsp_separator(self):
        assert parse_price("1\xa0299,99 zł") == 1299.99

    def test_pln_suffix(self):
        assert parse_price("450,00 PLN") == 450.0

    def test_none_input(self):
        assert parse_price(None) is None

    def test_empty_string(self):
        assert parse_price("") is None

    def test_below_threshold(self):
        assert parse_price("0.10") is None

    def test_plain_number(self):
        assert parse_price("89") == 89.0

    def test_thousands_dot(self):
        # European format: 1.299,99
        assert parse_price("1.299,99 zł") == 1299.99


# ── _extract_catalog_num ──────────────────────────────────────────────────────

class TestExtractCatalogNum:
    def test_standard_url(self):
        url = "https://mlamp.pl/pl/products/lampa-xeno-52408-saxby-srebna-85659"
        assert _extract_catalog_num(url) == "52408"

    def test_six_digit(self):
        url = "https://mlamp.pl/pl/products/oprawa-123456-led-white-99999"
        assert _extract_catalog_num(url) == "123456"

    def test_no_match(self):
        assert _extract_catalog_num("https://mlamp.pl/pl/products/lampa") is None

    def test_slug_only(self):
        assert _extract_catalog_num("lampa-xeno-52408-saxby") == "52408"


# ── _extract_product_id ───────────────────────────────────────────────────────

class TestExtractProductId:
    def test_numeric_url(self):
        assert _extract_product_id("https://www.ceneo.pl/123456789") == "123456789"

    def test_no_id(self):
        assert _extract_product_id("https://www.ceneo.pl/;szukaj-lampa") is None

    def test_short_number_ignored(self):
        assert _extract_product_id("https://www.ceneo.pl/123") is None


# ── _name_similarity ──────────────────────────────────────────────────────────

class TestNameSimilarity:
    def test_identical_names(self):
        name = "Lampa Xeno 52408 Saxby srebna"
        score = _name_similarity(name, name)
        assert score > 0.8

    def test_completely_different(self):
        score = _name_similarity("Lampa Xeno 52408", "Krzeslo biurowe ergonomiczne")
        assert score == 0.0

    def test_different_model_number_penalised(self):
        score_wrong = _name_similarity("Lampa Xeno 52408 Saxby", "Lampa Xeno 52409 Saxby")
        score_right = _name_similarity("Lampa Xeno 52408 Saxby", "Lampa Xeno 52408 Saxby")
        assert score_right > score_wrong

    def test_empty_names(self):
        assert _name_similarity("", "anything") == 0.0


# ── _model_anchors ────────────────────────────────────────────────────────────

class TestModelAnchors:
    def test_extracts_model_name_and_code(self):
        names, codes, ean = _model_anchors("Lampa Xeno 52408 Saxby srebna", "52408")
        assert "xeno" in names
        assert "52408" in codes

    def test_ean_passed_through(self):
        _, _, returned_ean = _model_anchors("Lampa Xeno 52408", "52408", ean="5901234567890")
        assert returned_ean == "5901234567890"

    def test_no_catalog_num(self):
        names, codes, _ = _model_anchors("Oprawa LED Slim Panel biala", None)
        # Should still find some distinctive words
        assert isinstance(names, set)
        assert isinstance(codes, set)


# ── _match_confidence ─────────────────────────────────────────────────────────

class TestMatchConfidence:
    def _anchors(self, product_name, catalog_num, ean=None):
        return _model_anchors(product_name, catalog_num, ean=ean)

    def test_full_match_returns_1(self):
        anchors = self._anchors("Lampa Xeno 52408 Saxby", "52408")
        assert _match_confidence(anchors, "Lampa Xeno 52408 LED zewnetrzna") == 1.0

    def test_ean_match_returns_1(self):
        anchors = self._anchors("Lampa Xeno 52408", "52408", ean="5901234567890")
        assert _match_confidence(anchors, "jakikolwiek tytul 5901234567890") == 1.0

    def test_code_only_returns_06(self):
        anchors = self._anchors("Lampa Xeno 52408 Saxby", "52408")
        conf = _match_confidence(anchors, "Zupelnie inne 52408 brak nazwy")
        assert conf == 0.6

    def test_no_match_returns_0(self):
        anchors = self._anchors("Lampa Xeno 52408 Saxby", "52408")
        assert _match_confidence(anchors, "Krzeslo biurowe ergonomiczne") == 0.0


# ── _build_search_queries ────────────────────────────────────────────────────

class TestBuildSearchQueries:
    def test_ean_is_first(self):
        queries = _build_search_queries("Lampa Xeno 52408 Saxby", "52408", ean="5901234567890")
        assert queries[0] == "5901234567890"

    def test_mfr_sku_second_when_different(self):
        queries = _build_search_queries("Lampa Xeno 52408 Saxby", "52408", mfr_sku="XEN-001")
        assert "XEN-001" in queries

    def test_no_duplicates(self):
        queries = _build_search_queries("Lampa Xeno 52408 Saxby", "52408")
        assert len(queries) == len(set(queries))

    def test_returns_list(self):
        queries = _build_search_queries("Lampa Xeno 52408 Saxby", "52408")
        assert isinstance(queries, list)
        assert len(queries) >= 1


# ── select_offers ─────────────────────────────────────────────────────────────

class TestSelectOffers:
    def _make_offers(self, prices):
        return [{"sklep": f"Sklep{i}", "cena": p, "url": f"http://x.pl/{i}"}
                for i, p in enumerate(prices)]

    def test_picks_cheaper_up_to_5(self):
        offers = self._make_offers([80, 90, 100, 110, 70, 85, 95])
        ref = 100.0
        selected, fallback = select_offers(offers, ref)
        assert not fallback
        assert all(o["cena"] < ref for o in selected)
        assert len(selected) <= 5

    def test_sorted_biggest_first(self):
        offers = self._make_offers([70, 80, 90])
        selected, _ = select_offers(offers, 100.0)
        prices = [o["cena"] for o in selected]
        assert prices == sorted(prices, reverse=True)

    def test_fallback_when_none_cheaper(self):
        offers = self._make_offers([110, 120, 130])
        selected, fallback = select_offers(offers, 100.0)
        assert fallback
        assert len(selected) <= 3

    def test_empty_offers(self):
        selected, fallback = select_offers([], 100.0)
        assert selected == []
        assert fallback is True

    def test_exact_ref_price_not_cheaper(self):
        offers = self._make_offers([100.0])
        selected, fallback = select_offers(offers, 100.0)
        assert fallback


# ── _extract_shop_name ────────────────────────────────────────────────────────

class TestExtractShopName:
    def _item(self, html):
        return BeautifulSoup(html, "lxml").body

    def test_img_alt(self):
        html = '<div class="product-offer"><img alt="MediaMarkt" src="logo.png"></div>'
        item = BeautifulSoup(html, "lxml").select_one(".product-offer")
        assert _extract_shop_name(item) == "MediaMarkt"

    def test_img_alt_skips_firma(self):
        html = '<div class="product-offer"><img alt="Firma" src="logo.png"><span class="shop-label">Neonet</span></div>'
        item = BeautifulSoup(html, "lxml").select_one(".product-offer")
        assert _extract_shop_name(item) == "Neonet"

    def test_fallback_default(self):
        html = '<div class="product-offer"><img alt="" src="logo.png"></div>'
        item = BeautifulSoup(html, "lxml").select_one(".product-offer")
        assert _extract_shop_name(item) == "Sklep"
