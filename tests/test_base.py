# coding: utf-8
"""test utils."""
from __future__ import unicode_literals


from unittest import TestCase

import xlsxwriter

from easyxlsx import FormatMixin


class FormatMixinTest(TestCase):
    """FormatMixin test."""

    def setUp(self):
        """setUp."""
        self.book = xlsxwriter.Workbook('test')

    def tearDown(self):
        """tearDown."""
        self.book.close()

    def test_formatmixin(self):
        """Test formatmixin."""
        formatmixn = FormatMixin(self.book)
