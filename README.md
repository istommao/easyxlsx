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

from easyxlsx import ModelExport


class UserExport(ModelExport):
    model = User
    fields = ('username', 'gender', 'age', 'email', 'added_at')


class UserAdmin(admin.ModelAdmin):

    actions = ['export_action']

    def export_action(self, request, queryset):
        data = UserExport(queryset).export()

        response = HttpResponse(data, content_type='application/vnd.ms-excel;charset=utf-8')
        response['Content-Disposition'] = 'attachment; filename="download.xls"'
        return response
```
