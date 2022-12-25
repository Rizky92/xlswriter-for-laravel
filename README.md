# About
Laravel wrapper package for php ext-xlswriter extension.  
  
[![Latest Stable Version](http://poser.pugx.org/rizky92/laravel-xlswriter/v)](https://packagist.org/packages/rizky92/laravel-xlswriter)
[![Total Downloads](http://poser.pugx.org/rizky92/laravel-xlswriter/downloads)](https://packagist.org/packages/rizky92/laravel-xlswriter)
[![License](http://poser.pugx.org/rizky92/laravel-xlswriter/license)](https://packagist.org/packages/rizky92/laravel-xlswriter)

## Requirements
- [laravel 7+](https://laravel.com/docs/7.x/installation)
- [ext-xlswriter 1.5.0+](https://pecl.php.net/package/xlswriter)

## Install
#### Via composer CLI
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
