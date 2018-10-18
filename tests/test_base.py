# coding: utf-8
"""test utils."""
from __future__ import unicode_literals


from unittest import TestCase

from easyxlsx import SimpleWriter


class SimpleWriterTest(TestCase):
    """FormatMixin test."""

    def setUp(self):
        self.headers = ('编号', '姓名', '年龄')

        dataset = (
            [1, '无声', 25],
            [2, '星尘', 26],
            [3, '黎明', 27],
        )
        self.dataset = dataset

    def test_export(self):
        writer = SimpleWriter(self.dataset, headers=self.headers)
        data = writer.export()

        self.assertTrue(isinstance(data, bytes))
        self.assertNotEqual(data, b'')

    def test_with_syntax(self):

        with SimpleWriter(self.dataset, headers=self.headers) as writer:
            data = writer.export()

        self.assertTrue(isinstance(data, bytes))
        self.assertNotEqual(data, b'')
