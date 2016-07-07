Yii2 Excel Export
=================

[![Latest Stable Version](https://poser.pugx.org/codemix/yii2-excelexport/v/stable)](https://packagist.org/packages/codemix/yii2-excelexport)
[![Total Downloads](https://poser.pugx.org/codemix/yii2-excelexport/downloads)](https://packagist.org/packages/codemix/yii2-excelexport)
[![Latest Unstable Version](https://poser.pugx.org/codemix/yii2-excelexport/v/unstable)](https://packagist.org/packages/codemix/yii2-excelexport)
[![License](https://poser.pugx.org/codemix/yii2-excelexport/license)](https://packagist.org/packages/codemix/yii2-excelexport)


## Features

 * Export data from `ActiveQuery` results
 * Export any other data (Array, Iterable, ...)
 * Create excel files with multiple sheets
 * Format cells and values

To write the Excel file, we use the excellent [PHPExcel](https://github.com/PHPOffice/PHPExcel) package.

## Installation

Install the package with [composer](http://getcomposer.org):

    composer require codemix/yii2-excelexport *

## Quickstart example

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

Find more examples below.


## Configuration and Use

### ExcelFile

Property | Description
---------|-------------
`writer` | The file format as supported by PHPOffice. The default is ` '\PHPExcel_Writer_Excel2007'`
`sheets` | An array of sheet configurations (see below). The keys are used as sheet names.

Methods | Description
---------|-------------
`saveAs($name)` | Saves the excel file under `$name`
`send($name=null,$inline=false)` | Sends the excel file to the browser. If `$name` is empty, the file is streamed for inline display, otherwhise a download dialog will open, unless `$inline` is `true` which will force inline display even if a filename is supplied.
`getWorkbook()` | Returns the `PHPExcel` workbook instance
`getTmpFile()` | Returns the `mikehaertl\tmp\File` instance of the temporary file

### ExcelSheet

Property | Description
---------|-------------
`data` | An array of data rows that should be used as sheet content
`titles` (optional) | An array of column titles
`types` (optional) | An array of types for specific columns as supported by PHPOffice, e.g. `\PHPExcel_Cell_DataType::TYPE_STRING`, indexed either by column name (e.g. `H`) or 0-based column index.
`formats` (optional) | An array of format strings for specific columns as supported by Excel, e.g. `#,##0.00`, indexed either by column name (e.g. `H`) or 0-based column index.
`formatters` (optional) | An array of value formatters for specific columns. Each must be a valid PHP callable whith the signature `function formatter($value, $row, $data)` where `$value` is the cell value to format, `$row` is the 0-based row index and `$data` is the current row data from the `data` configuration. The callbacks must be indexed either by column name (e.g. `H`) or by the 0-based column index.
`callbacks` (optional) | An array of callbacks for specific columns that should be called after rendering a cell, e.g. to apply some styling. Each must be a valid PHP callable with the signature `function callback($cell, $col, $row)` where `$cell` is the current `PHPExcel_Cell` object and `$col` and `$row` are the 0-based column and row indices respectively.

### ActiveExcelSheet

The class extends from `ExcelSheet` but differs in the following properties:

Property | Description
---------|-------------
`query` | The `ActiveQuery` for the row data (the `data` property will be ignored).
`data` | The read-only property that returns the batched query result.
`attributes` (optional) | The attributes to use as columns. Related attributes can be specifed in dot notation as usual, e.g. `team.name`. If not set, the `attributes()` from the corresponding `ActiveRecord` class will be used.
`titles` (optional) | The column titles, indexed by column name (e.g. `H`) or 0-based column index. If a column is not listed here, the respective attribute label will be used. If set to `false` no title row will be rendered.
`formats` (optional) | As in `ExcelSheet` but for `date`, `datetime` and `decimal` DB columns, the respective formats will be automatically set by default, according to the respective date format properties (see below) and the decimal precision.
`formatters` (optional) | As in `ExcelSheet` but for `date` and `datetime` columns the value will be autoconverted to the correct excel time format with `\PHPExcel_Shared_Date::PHPToExcel()` by default.
`dateFormat` | The excel format to use for `date` DB types. Default is `dd/mm/yyyy`.
`dateTimeFormat` | The excel format to use for `datetime` DB types. Default is `dd/mm/yyyy hh:mm:ss`.
`batchSize` | The query batchsize to use. Default is `100`.

## Examples

### ActiveQuery results

```php
<?php
$file = \Yii::createObject([
    'class' => 'codemix\excelexport\ExcelFile',

    'writer' => '\PHPExcel_Writer_Excel5', // Override default of `\PHPExcel_Writer_Excel2007`

    'sheets' => [

        'Active Users' => [
            'class' => 'codemix\excelexport\ActiveExcelSheet',
            'query' => User::find()->where(['active' => true]),

            // If not specified, all attributes from `User::attributes()` are used
            'attributes' => [
                'id',
                'name',
                'email',
                'team.name',    // Related attribute
                'created_at',
            ],

            // If not specified, the label from the respective record is used.
            // You can also override single titles, like here for the above `team.name`
            'titles' => [
                'D' => 'Team Name',
            ],
        ],

    ],
]);
$file->send('demo.xlsx');
```

### Raw data

```php
<?php
$file = \Yii::createObject([
    'class' => 'codemix\excelexport\ExcelFile',
    'sheets' => [

        'Result per Country' => [   // Name of the excel sheet
            'data' => [
                ['fr', 'France', 1.234, '2014-02-03 12:13:14'],
                ['de', 'Germany', 2.345, '2014-02-05 19:18:39'],
                ['uk', 'United Kingdom', 3.456, '2014-03-03 16:09:04'],
            ],

            // Set to `false` to suppress the title row
            'titles' => [
                'Code',
                'Name',
                'Volume',
                'Created At',
            ],

            'formats' => [
                // Either column name or 0-based column index can be used
                'C' => '#,##0.00',
                3 => 'dd/mm/yyyy hh:mm:ss',
            ],

            'formatters' => [
                // Dates and datetimes must be converted to Excel format
                3 => function ($value, $row, $data) {
                    return \PHPExcel_Shared_Date::PHPToExcel(strtotime($value));
                },
            ],
        ],

        'Countries' => [
            // Data for another sheet goes here ...
        ],
    ]
]);
// Save on disk
$file->saveAs('/tmp/export.xlsx');
```

### Query builder results

```php
<?php
$file = \Yii::createObject([
    'class' => 'codemix\excelexport\ExcelFile',
    'sheets' => [

        'Users' => [
            'data' => new (\yii\db\Query)
                ->select(['id','name','email'])
                ->from('user')
                ->each(100);
            'titles' => ['ID', 'Name', 'Email'],
        ],
    ]
]);
$file->send('demo.xlsx');
```

### Styling

As you have access to the `PHPExcel` object you can modify the excel file as you like.
For details please consult the
[PHPExcel documentation](https://github.com/PHPOffice/PHPExcel/tree/develop/Documentation/markdown/Overview).

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
$phpExcel = $file->getWorkbook();
$phpExcel->getSheet(1)->getStyle('B1')
    ->getFont()->getColor()->setARGB(\PHPExcel_Style_Color::COLOR_RED);
```

Alternatively you can use the callback feature from our `ExcelSheet`:

```php
<?php
$file = \Yii::createObject([
    'class' => 'codemix\excelexport\ExcelFile',
    'sheets' => [
        'Users' => [
            'class' => 'codemix\excelexport\ActiveExcelSheet',
            'query' => User::find(),
            'callbacks' => [
                // $cell is a PHPExcel_Cell object
                'A' => function ($cell, $row, $column) {
                    $cell->getStyle()->applyFromArray([
                        'font' => [
                            'bold' => true,
                        ],
                        'alignment' => [
                            'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
                        ],
                        'borders' => [
                            'top' => [
                                'style' => \PHPExcel_Style_Border::BORDER_THIN,
                            ],
                        ],
                        'fill' => [
                            'type' => \PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
                            'rotation' => 90,
                            'startcolor' => [
                                'argb' => 'FFA0A0A0',
                            ],
                            'endcolor' => [
                                'argb' => 'FFFFFFFF',
                            ],
                        ],
                    ]);
                },
            ],
        ],
    ],
]);

