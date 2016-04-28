# WIP This extension is still considered alpha stage!

Please use at your own risk - and leave feedback so that we can improve it! Thanks!

Yii2 Excel Export
=================

[![Latest Stable Version](https://poser.pugx.org/codemix/yii2-excelexport/v/stable.svg)](https://packagist.org/packages/codemix/yii2-excelexport)
[![Total Downloads](https://poser.pugx.org/codemix/yii2-excelexport/downloads)](https://packagist.org/packages/codemix/yii2-excelexport)
[![Latest Unstable Version](https://poser.pugx.org/codemix/yii2-excelexport/v/unstable.svg)](https://packagist.org/packages/codemix/yii2-excelexport)
[![License](https://poser.pugx.org/codemix/yii2-excelexport/license.svg)](https://packagist.org/packages/codemix/yii2-excelexport)


## Features

 * Export data from `ActiveQuery` results
 * Export any other data (Array, Iterable, ...)
 * Create excel files with multiple sheets
 * Format cells and values

> **Note:** To write the Excel file, we use the excellent
> [PHPExcel](https://github.com/PHPOffice/PHPExcel) package.

Here's a quick example to get you started:

```php
<?php
$file = \Yii::createObject([
    'class' => 'codemix\excelexport\ExcelFile',
    'sheets' => [
        'Users' => [
            'class' => 'codemix\excelexport\ActiveExcelSheet',
            'query' => User::find(),
        ]
    ]
]);
$file->send('user.xlsx');
```

## Installation

Install the package through [composer](http://getcomposer.org):

    composer require codemix/yii2-excelexport:1.0.0-alpha

Now you're ready to use the extension.


## Examples

For now we only provide some usage examples instead of a full documentation:

```php
<?php
$file = \Yii::createObject([
    'class' => 'codemix\excelexport\ExcelFile',
    'sheets' => [

        'Users' => [
            'class' => 'codemix\excelexport\ActiveExcelSheet',
            'query' => User::find(),
            'attributes' => [
                'id',
                'name',
                'email',
            ],

            // Default titles are attribute labels. To override a specific column,
            // index the title with the column index (0-based)
            'titles' => [
                1 => 'Full Name',
            ],
        ],

        'Countries' => [
            'data' => [
                ['fr', 'France', 1.234],
                ['de', 'Germany', 2.345],
                ['uk', 'United Kingdom', 3.456],
            ],
            'titles' => [
                'Code',
                'Name',
                'Volume'
            ],

            // Formats are set automatically for ActiveExcelSheets.
            'formats' => [
                2 => '#,##0.00',
            ],
        ],
    ]
]);
$file->send('demo.xlsx');
```
