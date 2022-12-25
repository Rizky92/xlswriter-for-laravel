## About
Laravel wrapper package for php ext-xlswriter extension

## Requirements
- laravel 7+
- ext-xlswriter 1.5.0+

## Install
### Via composer CLI
```
composer require rizky92/laravel-xlswriter
```

## Quick start
```php
use Rizky92\Xlswriter\ExcelExport;
use App\Models\User;

$users = User::all(['id', 'username', 'created_at']);
$columnHeaders = ['User ID', 'Username', 'Registration date'];

$excel = ExcelExport::make('users.xlsx', 'Sheet 1')
    ->setBasePath('excel/users')
    ->setDisk('public')
    ->setColumnHeaders($columnHeaders)
    ->setData($users)
    ->save();

return $excel->export();
```

## Documentation
TBA

## License
MIT