# coding: utf-8
"""easyxlsx base.py"""
import io
import datetime

import xlsxwriter


class FormatMixin(object):
    """FormatMixin."""

    header_font_size = 12
    body_font_size = 12

    def __init__(self, book):
        self.book = book
        self.headerfmt = None

        self.set_header_fmt({'font_size': self.header_font_size})

        self.bodyfmt = None

        self.set_body_fmt({'font_size': self.body_font_size})

    def set_header_fmt(self, fmt):
        """Set headerfmt."""
        self.headerfmt = self.book.add_format(fmt)
        self.headerfmt.set_bold()
        self.headerfmt.set_align('center')
        self.headerfmt.set_align('vcenter')

    def set_body_fmt(self, fmt):
        """Set bodyfmt."""
        self.bodyfmt = self.book.add_format(fmt)
        self.bodyfmt.set_align('center')
        self.bodyfmt.set_align('vcenter')


class BaseWriter(object):
    """BaseWriter."""

    datetime_style = 'yyyy-m-d hh:mm:ss'
    date_style = 'yyyy-m-d'
    rows_index = 0

    def __init__(self, sources, headers=None, bookname=None):
        """Init."""
        self.output = None

        if bookname:
            book = xlsxwriter.Workbook(bookname)
        else:
            self.output = io.BytesIO()
            book = xlsxwriter.Workbook(self.output, {'in_memory': True})

        self.header_line = headers
        self.book = book
        self.sources = sources

        self.formatmixin = FormatMixin(self.book)

    def export(self, sheet_name=None):
        """Export."""
        raise NotImplementedError('NotImplemented method export!')

    def get_export_data(self):
        """get_export_data."""
        self.close_book()

        if self.output:
            self.output.seek(0)
            iodata = self.output.read()

            self.output.close()
            return iodata

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close_book()

    def close_book(self):
        """Close."""
        self.book.close()


class SimpleWriter(BaseWriter):
    """Simple writer."""

    move_rows = None
    headers = ()

    def write_header(self, sheet):

        headerfmt = self.formatmixin.headerfmt

        headers = self.header_line or self.headers

        for col, value in enumerate(headers):
            sheet.write(self.rows_index, col, value, headerfmt)
        self.rows_index += 1

    def export(self, sheet_name=None):
        """Export."""
        sheet = self.book.add_worksheet(sheet_name)

        self.write_header(sheet)
        self.write_dataset(sheet, self.sources)

        return self.get_export_data()

    def get_fmt(self, value):
        """Get fmt."""
        if isinstance(value, datetime.datetime):
            datatimefmt = self.book.add_format(
                {'num_format': self.datetime_style})
            datatimefmt.set_align('center')
            datatimefmt.set_align('vcenter')
            return datatimefmt
        else:
            return self.formatmixin.bodyfmt

    def write_dataset(self, sheet, dataset):
        """Write dataset."""
        for linedata in dataset:
            self.move_rows = None
            for col, value in enumerate(linedata):
                sheet.write(self.rows_index, col, value, self.get_fmt(value))

            self.rows_index += self.move_rows if self.move_rows else 1


class StreamMixin(object):

    def get_export_data(self):
        """get_export_data."""
        self.close_book()

        if self.output:
            self.output.seek(0)
            iodata = self.output.read()

            self.output.close()

            yield iodata


class StreamSimpleWriter(SimpleWriter, StreamMixin):
    """Stream SimpleWriter."""


class ModelWriter(BaseWriter):
    """
    :功能描述: ModelWriter 根据 model导出
    """

    model = None
    fields = None
    move_rows = None
    headers = []
    fields_choices = []

    def write_header(self, sheet):
        """Write header."""
        model = self.model

        headerfmt = self.formatmixin.headerfmt

        for col, field in enumerate(self.fields):
            if field in self.headers:
                value = self.headers[field]
            else:
                value = model._meta.get_field(
                    field).verbose_name.title()    # pylint: disable=W0212

            sheet.write(self.rows_index, col, value, headerfmt)
        self.rows_index += 1

    def export(self, sheet_name=None):
        """Export."""
        sheet = self.book.add_worksheet(sheet_name)

        self.write_header(sheet)
        self.write_dataset(sheet, self.sources)

        return self.get_export_data()

    def get_fmt(self, value):
        """Get fmt."""
        if isinstance(value, datetime.datetime):
            datatimefmt = self.book.add_format(
                {'num_format': self.datetime_style})
            datatimefmt.set_align('center')
            datatimefmt.set_align('vcenter')
            return datatimefmt
        else:
            return self.formatmixin.bodyfmt

    def write_dataset(self, sheet, dataset):
        """Write dataset."""
        for obj in dataset:
            self.move_rows = None
            for col, field in enumerate(self.fields):

                if hasattr(self, field):
                    value = getattr(self, field)(obj)
                    sheet.write(self.rows_index, col, value,
                                self.get_fmt(value))
                else:
                    self.write_data(sheet, col, obj, field)

            self.rows_index += self.move_rows if self.move_rows else 1

    def write_data(self, sheet, col, obj, field):
        """Write data."""
        if '.' in field:
            attr, related_name = field.split('.')
            obj = getattr(obj, attr)
            value = getattr(obj, related_name)

            sheet.write(self.rows_index, col, value, self.get_fmt(value))
        elif hasattr(self, field):
            value = getattr(self, field)()
            sheet.write(self.rows_index, col, value, self.get_fmt(value))
        else:
            value = getattr(obj, field)
            if isinstance(value, datetime.datetime):
                value = value.replace(tzinfo=None)

            if field in self.fields_choices:
                value = self.fields_choices[field][value]

            sheet.write(self.rows_index, col, value, self.get_fmt(value))


class StreamModelWriter(StreamMixin, ModelWriter):
    """Stream Model writer."""
