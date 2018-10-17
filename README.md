# easyxlsx
easy way to use xlsxwriter


## Installation

```shell
pip install easyxlsx
```

## Usage

`Django`
```python
from django.http import HttpResponse

from easyxlsx import ModelWriter


class UserWriter(ModelWriter):
    model = User
    fields = ('username', 'gender', 'age', 'email', 'added_at')


class UserAdmin(admin.ModelAdmin):

    actions = ['writer_action']

    def writer_action(self, request, queryset):
        data = UserWriter(queryset).export()

        response = HttpResponse(data, content_type='application/vnd.ms-excel;charset=utf-8')
        response['Content-Disposition'] = 'attachment; filename="download.xls"'
        return response
```
